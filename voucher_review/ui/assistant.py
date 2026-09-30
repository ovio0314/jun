"""계정 질문하기 · 계정과목 현황 화면."""

from __future__ import annotations

import pandas as pd
import streamlit as st

from voucher.accounts import (
    SIDE_CREDIT,
    SIDE_DEBIT,
    code_name_conflicts,
    counterpart_accounts,
    file_overview,
    management_item_profile,
    with_side,
)
from voucher.export import to_excel_bytes
from voucher.qa import AccountAdvisor, list_examples

from .common import show_money


def render_assistant(result, inventory, guide: dict, ready: bool) -> None:
    def current_guide():
        return guide

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



def render_inventory(result, inventory, ready: bool, source_name: str) -> None:
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
                {"항목": "원본 파일", "값": source_name},
                {"항목": "시트/헤더행", "값": f"{result.sheet} / {result.header_row}"},
                {"항목": "차대구분", "값": "차변금액만 있으면 차변, 대변금액만 있으면 대변, 둘 다 있으면 차대동시"},
                {"항목": "빈 금액", "값": "0 으로 합산(상태는 빈값으로 보존)"},
                {"항목": "해석 불가 금액", "값": "0 대체 없이 금액오류로 분리, 합계 제외"},
                {"항목": "외화", "값": "원화 합계에 포함하지 않음"},
            ]),
        })
        st.download_button("계정과목 현황 엑셀 다운로드", xls, file_name="계정과목_현황.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

