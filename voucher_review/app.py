"""전표 계정과목 도우미 — Streamlit 진입점.

실행: streamlit run app.py   (기본 주소 http://127.0.0.1:8501)
"""

from __future__ import annotations

import json
from decimal import Decimal
from pathlib import Path

import pandas as pd
import streamlit as st

from voucher.accounts import (
    SIDE_CREDIT,
    SIDE_DEBIT,
    ZERO,
    account_inventory,
    code_name_conflicts,
    counterpart_accounts,
    file_overview,
    management_item_profile,
    voucher_balance,
    with_side,
)
from voucher.columns import EXPECTED_COLUMNS, REQUIRED_COLUMNS
from voucher.export import to_excel_bytes
from voucher.parser import VoucherFileError, auto_mapping, find_header_candidates, load_vouchers
from voucher.qa import AccountAdvisor, list_examples, load_guide, validate_guide

APP_DIR = Path(__file__).parent
INPUT_DIRS = [APP_DIR / "input", APP_DIR.parent / "input"]
LOCAL_GUIDE = APP_DIR / "local_data" / "account_guide.json"

st.set_page_config(page_title="전표 계정과목 도우미", page_icon="📒", layout="wide")


# ------------------------------------------------------------------ 유틸 ----
def won(v) -> str:
    if isinstance(v, Decimal):
        return f"{v:,.0f}" if v == v.to_integral_value() else f"{v:,}"
    return "" if v is None else str(v)


def show_money(df: pd.DataFrame, cols=("차변합계", "대변합계", "차이", "차변금액", "대변금액")) -> pd.DataFrame:
    out = df.copy()
    for c in cols:
        if c in out:
            out[c] = out[c].map(won)
    return out


def current_guide() -> dict:
    if "guide" not in st.session_state:
        st.session_state.guide = load_guide(LOCAL_GUIDE if LOCAL_GUIDE.exists() else None)
    return st.session_state.guide


def reset_session():
    keep = {"guide"}
    for k in list(st.session_state.keys()):
        if k not in keep:
            del st.session_state[k]


def input_files() -> list[Path]:
    files = []
    for d in INPUT_DIRS:
        if d.exists():
            files.extend(sorted(p for p in d.glob("*.xlsx") if not p.name.startswith("~$")))
    return files


# ------------------------------------------------------------------ 사이드바: 파일 ----
with st.sidebar:
    st.header("1. 전표 파일")
    st.caption("영림원 ERP 전표조건검색 엑셀(.xlsx). 파일은 이 PC 안에서만 처리하며 원본은 수정하지 않습니다.")
    uploaded = st.file_uploader("엑셀 업로드", type=["xlsx"])
    local = input_files()
    local_choice = None
    if local:
        local_choice = st.selectbox("또는 input 폴더 파일", ["(선택 안 함)"] + [p.name for p in local])

    source_bytes, source_name = None, None
    if uploaded is not None:
        source_bytes, source_name = uploaded.getvalue(), uploaded.name
    elif local_choice and local_choice != "(선택 안 함)":
        path = next(p for p in local if p.name == local_choice)
        source_bytes, source_name = path.read_bytes(), path.name

    if source_bytes is not None:
        try:
            candidates = find_header_candidates(source_bytes)
        except VoucherFileError as e:
            st.error(str(e))
            candidates = []
        if source_bytes and not candidates:
            st.error("전표 헤더(회계단위·기표번호·계정과목 등)가 있는 시트를 찾지 못했습니다.")
        if candidates:
            labels = [c.label for c in candidates]
            idx = st.selectbox("2. 시트 / 헤더 행 선택", range(len(labels)), format_func=lambda i: labels[i])
            cand = candidates[idx]
            with st.expander("3. 열 매핑 확인", expanded=False):
                st.caption("자동 인식된 열입니다. 헤더 이름이 다르면 직접 지정하세요. (*)는 필수")
                auto = auto_mapping(cand.headers)
                options = ["(없음)"] + [h for h in cand.headers if h]
                mapping = {}
                for col in EXPECTED_COLUMNS:
                    default = auto.get(col, "(없음)")
                    label = f"{col} *" if col in REQUIRED_COLUMNS else col
                    sel = st.selectbox(label, options, index=options.index(default) if default in options else 0,
                                       key=f"map_{col}")
                    if sel != "(없음)":
                        mapping[col] = sel
            if st.button("불러오기", type="primary", use_container_width=True):
                try:
                    res = load_vouchers(source_bytes, cand, mapping)
                except VoucherFileError as e:
                    st.error(str(e))
                else:
                    st.session_state.result = res
                    st.session_state.source_name = source_name
                    st.session_state.inventory = account_inventory(res.df) if not res.errors else None

    st.divider()
    if st.button("초기화 (현재 화면만 비우기)", use_container_width=True):
        reset_session()
        st.success("현재 세션을 비웠습니다. 원본 파일은 삭제하지 않았습니다.")
        st.rerun()

result = st.session_state.get("result")
inventory = st.session_state.get("inventory")
ready = result is not None and not result.errors and inventory is not None

st.title("📒 전표 계정과목 도우미")
st.caption(
    "1단계: 전표에 쓰인 차변·대변 계정과목 파악 + 제조업 전표 계정 질문 응답. "
    "답변은 참고용이며 회사 승인 계정 규칙이 아닙니다. 결재·반려·ERP 수정 기능은 없습니다."
)
if result is not None:
    for e in result.errors:
        st.error(e)
    for w in result.warnings:
        st.warning(w)
if ready:
    ov = file_overview(result.df)
    st.info(
        f"불러온 파일: **{st.session_state.source_name}** · 시트 {result.sheet} ({result.header_row}행 헤더"
        f"{', 제목: ' + result.title if result.title else ''}) · 회계일 {ov['회계일_시작']} ~ {ov['회계일_종료']} · "
        f"원본 {ov['원본행수']:,}행 / 전표 {ov['전표수']:,}건 / 계정과목 {ov['계정과목수']}개"
    )

tab_qa, tab_inv, tab_sum, tab_set = st.tabs(["💬 계정 질문하기", "📋 계정과목 현황", "🧾 파일 요약", "⚙️ 안내·설정"])

# ------------------------------------------------------------------ 질문하기 ----
with tab_qa:
    st.subheader("제조업 전표, 차변·대변에 어떤 계정을 쓰나요?")
    if not ready:
        st.caption("전표 파일을 불러오면 회사가 실제로 쓰는 계정코드와 과거 비슷한 적요의 처리 사례도 함께 보여 드립니다.")
    cols = st.columns(4)
    for i, ex in enumerate(list_examples()):
        if cols[i % 4].button(ex, key=f"ex_{i}", use_container_width=True):
            st.session_state.question = ex
    question = st.text_input("질문", key="question", placeholder="예: 공장 전기요금 낼 때 차변 대변 뭐 써요?")

    if question:
        advisor = AccountAdvisor(current_guide(), result.df if ready else None, inventory if ready else None)
        ans = advisor.ask(question)
        chips = []
        if ans.context:
            chips.append(f"비용 구분 추정: **{ans.context}**")
        if ans.side_focus:
            chips.append(f"질문한 쪽: **{ans.side_focus}**")
        if chips:
            st.caption(" · ".join(chips))
        for n in ans.notes:
            st.info(n)

        for rank, m in enumerate(ans.matches, start=1):
            s = m.scenario
            with st.container(border=True):
                st.markdown(f"### {rank}. {s['거래유형']}")
                st.caption(f"일치 키워드: {', '.join(m.matched_keywords)} · 가이드 버전 {current_guide().get('version', '')}")
                order = m.preferred_entries + [i for i in range(len(s["분개"])) if i not in m.preferred_entries]
                for i in order:
                    e = s["분개"][i]
                    star = " ⭐ 질문 맥락과 일치" if i in m.preferred_entries else ""
                    st.markdown(f"**▸ {e.get('조건', '')}**{star}")
                    c1, c2 = st.columns(2)
                    sides = [("차변", e["차변"], c1), ("대변", e["대변"], c2)]
                    if ans.side_focus == SIDE_CREDIT:
                        sides = [sides[1], sides[0]]
                    for side, items, col in sides:
                        col.markdown(f"**{side}**" + (" ← 질문" if ans.side_focus == side else ""))
                        for it in items:
                            col.markdown(f"- {it['계정']}" + (f" — _{it['비고']}_" if it.get("비고") else ""))
                if s.get("해설"):
                    st.markdown(f"📘 {s['해설']}")
                for c in s.get("주의", []):
                    st.markdown(f"⚠️ {c}")

                if ready:
                    names = []
                    for e in s["분개"]:
                        for it in e["차변"] + e["대변"]:
                            if it["계정"] in ans.company_accounts and it["계정"] not in names:
                                names.append(it["계정"])
                    rows = []
                    for n in names:
                        df = ans.company_accounts[n]
                        if df.empty:
                            rows.append({"가이드 계정": n, "일치구분": "회사 데이터에서 미발견"})
                        for r in df.to_dict("records"):
                            rows.append({"가이드 계정": n, **r})
                    st.markdown("**🏢 이 파일에서 회사가 실제로 쓴 계정코드**")
                    st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
                    st.caption("일치=계정명 동일, 동의어=가이드 동의어, 유사=계정명 포함. 코드 확정은 회계팀 기준을 따르세요.")

        if ready and ans.similar_memo is not None and len(ans.similar_memo):
            with st.container(border=True):
                st.markdown(f"### 🔎 과거 전표에서 비슷한 적요의 처리 사례")
                st.caption(f"적요 검색어: {', '.join(ans.similar_memo_terms)} · 과거 처리가 곧 정답은 아닙니다.")
                c1, c2 = st.columns(2)
                for side, col in ((SIDE_DEBIT, c1), (SIDE_CREDIT, c2)):
                    col.markdown(f"**{side}**")
                    col.dataframe(ans.similar_memo[ans.similar_memo["차대구분"] == side].drop(columns="차대구분"),
                                  hide_index=True, use_container_width=True)

        for prof in ans.mentioned_accounts:
            with st.container(border=True):
                st.markdown(f"### 📌 {prof['계정과목']} ({prof['계정과목코드']}) — 이 파일의 사용 현황")
                st.markdown(
                    f"사용구분 **{prof['사용구분']}** · 차변 {prof['차변행수']}행 / 대변 {prof['대변행수']}행 · "
                    f"사용 전표 {prof['사용전표수']}건 · 비용구분 {prof['비용구분'] or '-'}"
                )
                if prof.get("적요예시(상위)"):
                    st.caption(f"적요 예시: {prof['적요예시(상위)']}")
                c1, c2 = st.columns(2)
                c1.markdown("**차변에 쓸 때 대변 상대계정**")
                c1.dataframe(prof.get("차변일때_상대계정"), hide_index=True, use_container_width=True)
                c2.markdown("**대변에 쓸 때 차변 상대계정**")
                c2.dataframe(prof.get("대변일때_상대계정"), hide_index=True, use_container_width=True)

# ------------------------------------------------------------------ 계정과목 현황 ----
with tab_inv:
    if not ready:
        st.info("왼쪽에서 전표 파일을 불러오면 차변·대변에 쓰인 계정과목 전체 목록이 표시됩니다.")
    else:
        ov = file_overview(result.df)
        m = st.columns(4)
        m[0].metric("사용 계정과목", ov["계정과목수"])
        m[1].metric("차변에 쓰인 계정", ov["차변사용계정수"])
        m[2].metric("대변에 쓰인 계정", ov["대변사용계정수"])
        m[3].metric("양쪽 모두 쓰인 계정", int(((inventory["차변행수"] > 0) & (inventory["대변행수"] > 0)).sum()))

        f1, f2 = st.columns([1, 2])
        usage = f1.multiselect("사용구분", sorted(inventory["사용구분"].unique()), placeholder="전체")
        keyword = f2.text_input("계정코드/계정명/적요 검색")
        view = inventory
        if usage:
            view = view[view["사용구분"].isin(usage)]
        if keyword:
            k = keyword.strip()
            view = view[view["계정과목코드"].str.contains(k, regex=False)
                        | view["계정과목"].str.contains(k, regex=False)
                        | view["적요예시(상위)"].str.contains(k, regex=False)]
        st.dataframe(show_money(view), hide_index=True, use_container_width=True, height=420)

        conflicts = code_name_conflicts(result.df)
        if len(conflicts):
            st.warning("계정코드와 계정명이 1:1 이 아닌 경우가 있습니다. 확인이 필요합니다.")
            st.dataframe(conflicts, hide_index=True, use_container_width=True)

        st.markdown("#### 계정 상세")
        labels = [f"{r.계정과목코드} {r.계정과목}" for r in inventory.itertuples()]
        pick = st.selectbox("계정 선택", range(len(labels)), format_func=lambda i: labels[i])
        row = inventory.iloc[pick]
        code = row["계정과목코드"]
        c1, c2 = st.columns(2)
        c1.markdown("**차변으로 쓴 전표의 대변 상대계정**")
        c1.dataframe(counterpart_accounts(result.df, code, SIDE_DEBIT), hide_index=True, use_container_width=True)
        c2.markdown("**대변으로 쓴 전표의 차변 상대계정**")
        c2.dataframe(counterpart_accounts(result.df, code, SIDE_CREDIT), hide_index=True, use_container_width=True)
        st.markdown("**관리항목 사용 현황** (같은 관리항목 열이라도 계정마다 의미가 다릅니다)")
        st.dataframe(management_item_profile(result.df, code), hide_index=True, use_container_width=True)
        lines = with_side(result.df[result.df["계정과목코드"] == code])
        st.markdown("**분개 행 (원본 엑셀 행번호 포함)**")
        st.dataframe(
            show_money(lines[["_엑셀행", "기표번호", "행번호", "회계일", "기표부서", "기표자", "차대구분",
                              "차변금액", "대변금액", "적요", "비용구분"]]),
            hide_index=True, use_container_width=True, height=300,
        )

        xls = to_excel_bytes({
            "계정과목현황": inventory,
            "코드명불일치": conflicts,
            "적용기준": pd.DataFrame([
                {"항목": "원본 파일", "값": st.session_state.source_name},
                {"항목": "시트/헤더행", "값": f"{result.sheet} / {result.header_row}"},
                {"항목": "차대구분", "값": "차변금액만 있으면 차변, 대변금액만 있으면 대변, 둘 다 있으면 차대동시"},
                {"항목": "빈 금액", "값": "0 으로 합산(상태는 빈값으로 보존)"},
                {"항목": "해석 불가 금액", "값": "0 대체 없이 금액오류로 분리, 합계 제외"},
                {"항목": "외화", "값": "원화 합계에 포함하지 않음"},
            ]),
        })
        st.download_button("계정과목 현황 엑셀 다운로드", xls, file_name="계정과목_현황.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ------------------------------------------------------------------ 파일 요약 ----
with tab_sum:
    if not ready:
        st.info("전표 파일을 불러오면 요약이 표시됩니다.")
    else:
        ov = file_overview(result.df)
        m = st.columns(5)
        m[0].metric("원본 행 수", f"{ov['원본행수']:,}")
        m[1].metric("전표 수 (회계단위+전표관리단위+기표번호)", f"{ov['전표수']:,}")
        m[2].metric("기표번호 수", f"{ov['기표번호수']:,}")
        m[3].metric("차대변 불일치 또는 추출 누락 확인", ov["차대변불일치전표수"])
        m[4].metric("적요 빈칸 행", ov["적요빈칸행수"])
        st.caption(
            f"대상 기간(파일 기준): {ov['회계일_시작']} ~ {ov['회계일_종료']}. 이 파일의 전표 수를 업무상 검토 건수"
            "(약 750~800건, 집계 기간 미확인)와 같은 모집단으로 가정하지 않습니다."
        )
        st.markdown("**파일첨부여부 값 구분** — FALSE 여도 '증빙 누락'으로 판정하지 않습니다 (ERP 첨부 표시일 뿐, 다우오피스 증빙 미확인).")
        st.write(ov["파일첨부여부구분"])
        bal = voucher_balance(result.df)
        diff = bal[bal["차이"] != ZERO]
        if len(diff):
            st.warning("차대변 불일치 또는 추출 누락 확인 대상 (필터된 불완전 전표일 수 있습니다)")
            st.dataframe(show_money(diff), hide_index=True, use_container_width=True)

# ------------------------------------------------------------------ 안내·설정 ----
with tab_set:
    guide = current_guide()
    st.markdown(f"**계정 가이드 버전:** {guide.get('version', '-')}")
    st.caption(guide.get("기준", ""))
    st.markdown(
        """
**처리 기준 (보수적 기본값)**
- 전표 키 = 회계단위 + 전표관리단위 + 기표번호. 전표기표번호는 분개 행 식별자로만 사용합니다.
- 계정과목코드·번호·사업자번호는 문자열로 보존하고, 공백 정리 전 원본 값도 따로 보관합니다.
- 금액: 쉼표·공백·'원' 제거, 음수(-, △, 괄호) 허용. 해석 불가 값은 0 으로 바꾸지 않고 **금액오류**로 분리합니다.
- 빈 금액은 합계에서 0 으로 봅니다(ERP 는 한쪽만 입력). 빈값 여부는 상태 열에 남깁니다.
- 외화 금액은 원화 합계에 더하지 않습니다.
- 파일첨부여부는 문자열 TRUE/FALSE, 실제 불리언, 빈값을 구분해 표시만 하고 증빙 판정에 쓰지 않습니다.
- 질문 답변의 분개는 **일반기업회계기준 제조업 관행 참고**이며, 회사 승인 규칙이 아닙니다. 거래처만으로 계정을 확정하지 않습니다.
        """
    )
    with st.expander("가이드 거래유형 목록 보기"):
        st.dataframe(pd.DataFrame([{"ID": s["id"], "거래유형": s["거래유형"], "키워드": ", ".join(s["키워드"])}
                                   for s in guide["scenarios"]]), hide_index=True, use_container_width=True)
    st.download_button("현재 가이드 JSON 다운로드", json.dumps(guide, ensure_ascii=False, indent=2).encode("utf-8"),
                       file_name="account_guide.json", mime="application/json")
    up = st.file_uploader("회사 기준으로 수정한 가이드 JSON 올리기", type=["json"], key="guide_upload")
    if up is not None:
        try:
            new_guide = json.loads(up.getvalue().decode("utf-8"))
            validate_guide(new_guide)
        except (ValueError, UnicodeDecodeError) as e:
            st.error(f"가이드 JSON 을 읽을 수 없습니다: {e}")
        else:
            c1, c2 = st.columns(2)
            if c1.button("이번 세션에만 적용"):
                st.session_state.guide = new_guide
                st.success("적용했습니다.")
            if c2.button("이 PC에 저장 (local_data/account_guide.json)"):
                LOCAL_GUIDE.parent.mkdir(exist_ok=True)
                LOCAL_GUIDE.write_text(json.dumps(new_guide, ensure_ascii=False, indent=2), encoding="utf-8")
                st.session_state.guide = new_guide
                st.success("저장했습니다.")
    if LOCAL_GUIDE.exists():
        st.markdown("---")
        st.markdown(f"저장된 회사 가이드: `{LOCAL_GUIDE.relative_to(APP_DIR)}`")
        confirm = st.checkbox("저장된 회사 가이드 파일을 삭제하고 기본 가이드로 되돌리는 것에 동의합니다.")
        if st.button("저장된 가이드 삭제", disabled=not confirm):
            LOCAL_GUIDE.unlink()
            st.session_state.guide = load_guide()
            st.success("삭제했습니다. 기본 가이드를 사용합니다.")
