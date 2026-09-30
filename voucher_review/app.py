"""전표 검토 도우미 — Streamlit 진입점.

실행: streamlit run app.py   (기본 주소 http://127.0.0.1:8501, 로컬 전용)
화면 코드는 ui/, 파서·규칙 엔진은 voucher/ 에 있다.
"""

from __future__ import annotations

import hashlib
import os
from pathlib import Path

import streamlit as st

from ui.assistant import render_assistant, render_inventory
from ui.common import flash_and_rerun, show_flash
from ui.review_page import render_review
from ui.settings_page import render_settings
from ui.state import AppState
from ui.storage_page import render_storage
from voucher.accounts import file_overview
from voucher.columns import EXPECTED_COLUMNS, REQUIRED_COLUMNS
from voucher.parser import VoucherFileError, auto_mapping, find_header_candidates, load_vouchers
from voucher.qa import load_guide
from voucher.settings import load_settings

APP_DIR = Path(__file__).parent
# 환경변수로 바꿀 수 있음 (테스트 격리용)
INPUT_DIRS = [Path(p) for p in os.environ.get("VOUCHER_INPUT_DIRS", "").split(os.pathsep) if p] or \
    [APP_DIR / "input", APP_DIR.parent / "input"]
LOCAL_DIR = Path(os.environ.get("VOUCHER_LOCAL_DIR") or APP_DIR / "local_data")
SETTINGS_PATH = LOCAL_DIR / "settings.json"
LOCAL_GUIDE = LOCAL_DIR / "account_guide.json"
SAVE_DIR = LOCAL_DIR / "saves"

st.set_page_config(page_title="전표 검토 도우미", page_icon="📒", layout="wide")
state = AppState(lambda: load_settings(SETTINGS_PATH))
if "guide" not in st.session_state:
    st.session_state.guide = load_guide(LOCAL_GUIDE if LOCAL_GUIDE.exists() else None)


def input_files() -> list[Path]:
    files = []
    for d in INPUT_DIRS:
        if d.exists():
            files.extend(sorted(p for p in d.glob("*.xlsx") if not p.name.startswith("~$")))
    return files


# ------------------------------------------------------------------ 사이드바 ----
with st.sidebar:
    st.header("1. 전표 파일")
    st.caption("영림원 ERP 전표조건검색 엑셀(.xlsx). 이 PC 안에서만 처리하며 원본은 수정하지 않습니다.")
    uploaded = st.file_uploader("엑셀 업로드", type=["xlsx"])
    local = input_files()
    local_choice = st.selectbox("또는 input 폴더 파일", ["(선택 안 함)"] + [p.name for p in local]) if local else None

    source_bytes, source_name = None, None
    if uploaded is not None:
        source_bytes, source_name = uploaded.getvalue(), uploaded.name
    elif local_choice and local_choice != "(선택 안 함)":
        source_bytes = next(p for p in local if p.name == local_choice).read_bytes()
        source_name = local_choice

    if source_bytes is not None:
        try:
            candidates = find_header_candidates(source_bytes)
        except VoucherFileError as e:
            st.error(str(e))
            candidates = []
        else:
            if not candidates:
                st.error("전표 헤더(회계단위·기표번호·계정과목 등)가 있는 시트를 찾지 못했습니다.")
        if candidates:
            labels = [c.label for c in candidates]
            idx = st.selectbox("2. 시트 / 헤더 행 선택", range(len(labels)), format_func=lambda i: labels[i])
            cand = candidates[idx]
            if cand.preamble:
                st.caption("헤더 위 내용: " + " | ".join(cand.preamble))
            with st.expander("3. 열 매핑 확인"):
                st.caption("자동 인식된 열입니다. 헤더 이름이 다르면 직접 지정하세요. (*) 필수")
                auto = auto_mapping(cand.headers)
                options = ["(없음)"] + [h for h in cand.headers if h]
                mapping = {}
                for col in EXPECTED_COLUMNS:
                    default = auto.get(col, "(없음)")
                    sel = st.selectbox(f"{col} *" if col in REQUIRED_COLUMNS else col, options,
                                       index=options.index(default) if default in options else 0, key=f"map_{col}")
                    if sel != "(없음)":
                        mapping[col] = sel
            if st.button("불러오기", use_container_width=True):
                try:
                    res = load_vouchers(source_bytes, cand, mapping)
                except VoucherFileError as e:
                    st.error(str(e))
                else:
                    state.set_loaded(res, source_name, hashlib.sha256(source_bytes).hexdigest())
                    flash_and_rerun(f"불러왔습니다: {source_name}. '검사 실행'을 눌러 주세요.", "info")

    st.header("4. 검사")
    loaded_ok = state.result is not None and not state.result.errors
    if st.button("검사 실행", type="primary", use_container_width=True, disabled=not loaded_ok):
        state.rerun_checks()
        flash_and_rerun("검사를 실행했습니다. 검토상태는 바뀌지 않습니다.")
    state.reviewer = st.text_input("검토자 이름", value=state.reviewer)
    st.divider()
    if st.button("초기화 (현재 세션만 비우기)", use_container_width=True):
        state.reset()
        flash_and_rerun("현재 세션을 비웠습니다. 원본 파일과 저장 데이터는 삭제하지 않았습니다.")

# ------------------------------------------------------------------ 본문 ----
result = state.result
ready = result is not None and not result.errors and state.inventory is not None

st.title("📒 전표 검토 도우미")
st.caption("전표별 자동 검사 결과와 계정 질문 응답. 결재·반려·자동 분개 수정·ERP 쓰기 기능은 없습니다. "
           "검사 통과는 설정된 기준에 대한 결과이며 최종 회계 적정성 보증이 아닙니다.")
show_flash()
if result is not None:
    for e in result.errors:
        st.error(e)
    for w in result.warnings:
        st.warning(w)
if ready:
    ov = file_overview(result.df)
    st.info(f"파일: **{state.source_name}** · 시트 {result.sheet} ({result.header_row}행 헤더) · "
            f"회계일 {ov['회계일_시작']} ~ {ov['회계일_종료']} · 원본 {ov['원본행수']:,}행 / 전표 {ov['전표수']:,}건")

tabs = st.tabs(["🔍 전표 검사", "💬 계정 질문하기", "📋 계정과목 현황", "⚙️ 검사 설정", "💾 저장·불러오기"])
with tabs[0]:
    render_review(state)
with tabs[1]:
    render_assistant(result, state.inventory, st.session_state.guide, ready)
with tabs[2]:
    render_inventory(result, state.inventory, ready, state.source_name)
with tabs[3]:
    render_settings(state, SETTINGS_PATH, LOCAL_GUIDE)
with tabs[4]:
    render_storage(state, SAVE_DIR)
