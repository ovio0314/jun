"""로컬 저장 · 재열기 · 검토상태 이어받기 · 저장 데이터 삭제."""

from __future__ import annotations

import datetime as dt

import streamlit as st

from voucher.evidence import ManualEvidenceStore
from voucher.parser import LoadResult
from voucher.review import (
    content_hashes,
    delete_save,
    init_reviews,
    list_saves,
    load_session,
    safe_save_name,
    save_session,
)
from voucher.settings import merge_defaults

from .common import flash_and_rerun


def render_storage(state, save_dir) -> None:
    st.caption(f"저장 위치: `{save_dir}` (이 PC 로컬, 저장소에 올리지 않음). 저장 파일에는 전표 데이터·검토 메모·수기 증빙이 들어 있습니다.")
    res = state.result
    if res is not None and not res.errors:
        st.subheader("현재 작업 저장")
        name = st.text_input("저장 이름", value=f"{dt.date.today():%Y%m%d}_{state.source_name.rsplit('.', 1)[0]}")
        if st.button("저장", type="primary"):
            path = save_session(save_dir / f"{safe_save_name(name)}.json", source_name=state.source_name,
                                source_sha256=state.source_sha256, sheet=res.sheet, header_row=res.header_row,
                                raw=res.raw, df=res.df, reviews=state.reviews, evidence_json=state.evidence.dumps(),
                                settings=state.settings, preamble=res.preamble)
            flash_and_rerun(f"저장했습니다: {path.name}")

    saves = list_saves(save_dir)
    if not saves:
        st.info("저장된 작업이 없습니다.")
        return
    st.subheader("저장된 작업")
    pick = st.selectbox("저장 파일", saves, format_func=lambda p: f"{p.stem}  ({dt.datetime.fromtimestamp(p.stat().st_mtime):%Y-%m-%d %H:%M})")
    c = st.columns(2)
    if c[0].button("재열기 (저장 당시 데이터로)"):
        data = load_session(pick)
        loaded = LoadResult(data["raw_df"], data["df"], {}, data["source"]["sheet"], data["source"]["header_row"], "",
                            preamble=data["source"].get("preamble", []))
        state.set_loaded(loaded, data["source"]["name"], data["source"]["sha256"], data["review_states"])
        state.evidence = ManualEvidenceStore.from_dict(data["evidence"])
        state.settings = merge_defaults(data["settings"])
        state.rerun_checks()
        flash_and_rerun(f"재열었습니다: {pick.stem} (저장 당시 설정으로 다시 검사)")
    if c[1].button("현재 불러온 파일에 검토상태·증빙 이어받기", disabled=res is None or bool(res.errors)):
        data = load_session(pick)
        state.reviews = init_reviews(content_hashes(res.df), data["review_states"])
        store = ManualEvidenceStore.from_dict(data["evidence"])
        keys = set(res.df["전표키"])
        store.links = [l for l in store.links if l.voucher_key in keys]
        state.evidence = store
        state.rerun_checks()
        changed = sum(1 for r in state.reviews.values() if r.변경감지)
        flash_and_rerun(f"이어받았습니다. 분개가 바뀐 전표 {changed}건은 검토완료를 승계하지 않고 '미검토'로 표시했습니다.")

    st.divider()
    st.subheader("저장 데이터 삭제")
    st.caption("원본 엑셀 파일은 삭제하지 않습니다. 선택한 저장 파일 1개만 삭제합니다.")
    target = st.selectbox("삭제할 저장 파일", saves, format_func=lambda p: p.name, key="del_target")
    typed = st.text_input("확인을 위해 삭제할 파일 이름(확장자 제외)을 그대로 입력하세요")
    if st.button("저장 파일 삭제", disabled=typed != target.stem):
        delete_save(target, save_dir)
        flash_and_rerun(f"삭제했습니다: {target.name}")
