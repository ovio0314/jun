"""전표 검사 화면: 요약 · 필터 · 목록 · 상세 · 검토 메모 · 증빙 수기 연결 · 결과 다운로드."""

from __future__ import annotations

import datetime as dt
from decimal import Decimal

import pandas as pd
import streamlit as st

from voucher.checks import AREAS, SUMMARY_ORDER, sort_error_first, summary_counts
from voucher.columns import MANAGEMENT_ITEM_COLUMNS
from voucher.evidence import APPROVAL_DOC, BY_MANUAL, DOC_FIELDS, TAX_INVOICE
from voucher.export import build_result_sheets, check_export_counts, to_excel_bytes
from voucher.parser import parse_amount
from voucher.review import REVIEW_STATES, reviews_frame, update_review
from voucher.settings import settings_version

from .common import flash_and_rerun, icons, show_money


def apply_filters(summary: pd.DataFrame, reviews: pd.DataFrame, f: dict) -> pd.DataFrame:
    """전표 단위 필터 (행 단위가 아니므로 중복 집계 없음)."""
    v = summary.merge(reviews[["전표키", "검토상태", "변경감지"]], on="전표키", how="left", validate="one_to_one")
    v["검토상태"] = v["검토상태"].fillna("미검토")
    if f.get("date"):
        start, end = f["date"]
        v = v[v["회계일"].map(lambda d: d is not None and start <= d <= end)]
    for col in ("기표부서", "기표자", "결재상태", "전표유형", "검사결과", "검토상태"):
        if f.get(col):
            v = v[v[col].isin(f[col])]
    return v


def _criteria_frame(state) -> pd.DataFrame:
    s = state.settings
    res = state.result
    rows = [
        ("원본 파일", state.source_name), ("원본 SHA-256", state.source_sha256[:16]),
        ("시트/헤더행", f"{res.sheet} / {res.header_row}"), ("헤더 위 조회 조건", " | ".join(res.preamble)),
        ("설정 버전", settings_version(s)),
        ("검사 실행 시각", state.checked_at),
        ("요약 기준", "오류 > 확인 필요 > 미검사 > 설정된 검사 통과 (검사 통과는 최종 회계 적정성 보증 아님)"),
        ("부가세 계정코드", ", ".join(s["vat"]["account_codes"])), ("부가세 계정명 키워드", ", ".join(s["vat"]["name_keywords"])),
        ("공급가액 열 매핑", ", ".join(f"{k}→{v}" for k, v in s["vat"]["supply_columns"].items()) or "미설정"),
        ("일반과세 증빙 유형", ", ".join(s["vat"]["general_evidence_values"]) or "미설정"),
        ("특수 증빙 제외 키워드", ", ".join(s["vat"]["excluded_evidence_keywords"])),
        ("반올림 기준", s["vat"]["rounding"] or "미설정"), ("허용 차이(원)", s["vat"]["tolerance"]),
        ("적요 빈칸 상태", s["memo"]["blank_status"]),
        ("증빙 요구", f"부가세 전표 세금계산서={s['evidence']['require_tax_invoice_for_vat']}, 품의서={s['evidence']['require_approval_doc']}"),
        ("파일첨부여부", "증빙 판정에 사용하지 않음 (ERP 첨부 표시 FALSE, 다우오피스 증빙 미확인)"),
    ]
    for r in s.get("account_rules", []):
        rows.append((f"계정규칙 {r.get('규칙ID')}@v{r.get('버전')}", f"활성={r.get('활성')} · {r.get('이름')} · 근거: {r.get('근거')}"))
    for r in s["memo"].get("rules", []):
        rows.append((f"적요규칙 {r.get('규칙ID')}@v{r.get('버전')}", f"활성={r.get('활성')} · {r.get('이름')} · 근거: {r.get('근거')}"))
    return pd.DataFrame(rows, columns=["항목", "값"])


def render_review(state) -> None:
    res = state.result
    if res is None or res.errors:
        st.info("왼쪽에서 전표 파일을 불러온 뒤 **검사 실행**을 눌러 주세요.")
        return
    if state.summary is None:
        st.info("파일을 불러왔습니다. 왼쪽의 **검사 실행**을 눌러 주세요.")
        return
    if state.checked_version != settings_version(state.settings):
        st.warning("검사 설정이 바뀌었습니다. 최신 설정으로 보려면 **검사 실행**을 다시 눌러 주세요.")

    summary, results = state.summary, state.results
    reviews = reviews_frame(state.reviews)
    counts = summary_counts(summary)
    df = res.df

    dates = [d for d in df["회계일"] if d is not None]
    st.caption(
        f"대상 기간(파일 기준): {min(dates) if dates else '-'} ~ {max(dates) if dates else '-'} · "
        f"헤더 위 조회 조건: {' | '.join(res.preamble) or '(없음)'} · 설정 버전 {state.checked_version}. "
        "이 파일의 전표 수를 업무상 검토 건수(약 750~800건, 집계 기간 미확인)와 같은 모집단으로 가정하지 않습니다."
    )
    m = st.columns(6)
    m[0].metric("원본 행 수", f"{len(df):,}")
    m[1].metric("전표 수", f"{counts['전체']:,}")
    m[2].metric("🔴 오류", counts["오류"])
    m[3].metric("🟠 확인 필요", counts["확인 필요"])
    m[4].metric("⚪ 미검사", counts["미검사"])
    m[5].metric("🟢 설정된 검사 통과", counts["설정된 검사 통과"])
    changed = int(reviews["변경감지"].sum()) if len(reviews) else 0
    if changed:
        st.warning(f"이전 저장 이후 분개 내용이 바뀐 전표 {changed}건 — 기존 검토완료를 승계하지 않고 '미검토'로 되돌렸습니다.")

    # ---- 필터 (초기값 전체)
    with st.expander("필터", expanded=True):
        c = st.columns(4)
        f = {}
        if dates:
            picked = c[0].date_input("회계일", (min(dates), max(dates)), min_value=min(dates), max_value=max(dates))
            # 전체 기간 그대로면 필터하지 않음 (회계일이 빈 전표도 숨기지 않기 위해)
            if isinstance(picked, (tuple, list)) and len(picked) == 2 and tuple(picked) != (min(dates), max(dates)):
                f["date"] = picked
        for i, col in enumerate(["기표부서", "기표자", "결재상태", "전표유형"]):
            f[col] = c[(i + 1) % 4].multiselect(col, sorted(v for v in summary[col].unique() if v), placeholder="전체")
        c2 = st.columns(3)
        f["검사결과"] = c2[0].multiselect("검사결과", SUMMARY_ORDER, placeholder="전체")
        f["검토상태"] = c2[1].multiselect("검토상태", REVIEW_STATES, placeholder="전체")
        order = c2[2].selectbox("정렬", ["오류 우선", "회계일", "기표번호"])
    view = apply_filters(summary, reviews, f)
    if order == "오류 우선":
        view = sort_error_first(view)
    else:
        view = view.sort_values(order, na_position="last")
    fc = {s: int((view["검사결과"] == s).sum()) for s in SUMMARY_ORDER}
    st.caption(f"필터 결과 {len(view)}건 = " + " + ".join(f"{k} {v}" for k, v in fc.items()) + " (미검사 전표도 숨기지 않음)")

    cols = ["기표번호", "회계일", "기표자", "기표부서", "대표적요", "차변합계", "대변합계", "차이", *AREAS, "검사결과", "사유",
            "검토상태", "변경감지", "회계단위", "전표관리단위"]
    st.dataframe(icons(show_money(view[cols]), [*AREAS, "검사결과"]), hide_index=True, use_container_width=True, height=380)

    # ---- 다운로드
    sheets = build_result_sheets(summary, results, res.raw, df, reviews, _criteria_frame(state))
    problems = check_export_counts(sheets, len(df), counts["전체"])
    if problems:
        st.error("결과 건수 대조 실패: " + "; ".join(problems))
    else:
        st.caption(f"다운로드 건수 대조: 원본분개 {len(sheets['원본분개'])}행 = 원본 {len(df)}행, 전표요약 {len(sheets['전표요약'])}건 = 전표 {counts['전체']}건")
    st.download_button("검사 결과 엑셀 다운로드 (전표요약·검사상세·원본분개·적용기준)", to_excel_bytes(sheets),
                       file_name=f"전표검사결과_{dt.date.today():%Y%m%d}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ---- 상세
    st.markdown("### 전표 상세")
    if view.empty:
        st.info("필터 조건에 맞는 전표가 없습니다.")
        return
    keys = list(view["전표키"])
    label = {r["전표키"]: f"{r['기표번호']} · {r['검사결과']} · {r['대표적요'][:30]}" for r in view.to_dict("records")}
    key = st.selectbox("전표 선택", keys, format_func=lambda k: label[k])
    render_detail(state, key)


def render_detail(state, key: str) -> None:
    df = state.result.df
    lines = df[df["전표키"] == key]
    srow = state.summary[state.summary["전표키"] == key].iloc[0]
    st.markdown(f"**전표 키** `{key}` · 요약 **{srow['검사결과']}** · " +
                " · ".join(f"{a} {srow[a]}" for a in AREAS))
    st.caption("검사 통과는 설정된 기준에 대한 결과이며 최종 회계 적정성 보증이 아닙니다.")
    show = ["_엑셀행", "전표기표번호", "행번호", "계정과목코드", "계정과목", "차변금액", "대변금액", "적요", "증빙",
            "거래처사업자번호", "비용구분", "귀속부서", "파일첨부여부", "파일첨부여부_구분", *MANAGEMENT_ITEM_COLUMNS]
    st.markdown("**분개 전체 (관리항목은 원본 값 그대로)**")
    raw = state.result.raw.loc[lines.index]
    table = show_money(lines[[c for c in show if c in lines.columns]])
    for c in MANAGEMENT_ITEM_COLUMNS:
        table[c] = raw[c].values
    st.dataframe(table, hide_index=True, use_container_width=True)

    st.markdown("**검사 근거·해설**")
    r = state.results[state.results["전표키"] == key]
    st.dataframe(icons(r.drop(columns="전표키"), ["상태"]), hide_index=True, use_container_width=True)

    c1, c2 = st.columns(2)
    with c1, st.form(f"review_{key}"):
        rv = state.reviews[key]
        st.markdown("**검토 (자동 검사와 별도로 저장)**")
        if rv.변경감지:
            st.warning(f"분개 내용 변경 감지 — 이전 검토상태 '{rv.이전검토상태}' 는 승계하지 않았습니다.")
        status = st.selectbox("검토상태", REVIEW_STATES, index=REVIEW_STATES.index(rv.검토상태))
        reviewer = st.text_input("검토자", value=rv.검토자 or state.reviewer)
        memo = st.text_area("메모", value=rv.메모)
        st.caption(f"마지막 검토일시: {rv.검토일시 or '-'}")
        if st.form_submit_button("검토 저장"):
            update_review(state.reviews, key, status=status, memo=memo, reviewer=reviewer)
            flash_and_rerun("검토 내용을 반영했습니다. (파일로 남기려면 '저장·불러오기' 탭에서 저장)")
    with c2:
        render_evidence_panel(state, key)


def render_evidence_panel(state, key: str) -> None:
    store = state.evidence
    st.markdown("**증빙 연결 (수기 입력 · 2차에서 다우오피스/OCR 연동 예정)**")
    st.caption("연결은 전표 키 기준 수동 연결입니다. 파일명·ERP 첨부 표시·'증빙' 열 텍스트로 적정성을 판단하지 않습니다.")
    for link in store.links_for(key):
        doc = store.get_doc(link.doc_id)
        desc = f"{doc.doc_type} · " + ", ".join(f"{k}={doc.fields[k].value}" for k in DOC_FIELDS[doc.doc_type][:3]
                                                if k in doc.fields) if doc else "(문서 없음)"
        cc = st.columns([4, 1])
        cc[0].write(f"[{link.status}/{link.method}] {desc}")
        if cc[1].button("해제", key=f"unlink_{key}_{link.doc_id}"):
            store.unlink(key, link.doc_id)
            state.rerun_checks()
            flash_and_rerun("증빙 연결을 해제하고 다시 검사했습니다.")
    kind = st.radio("증빙 유형", [TAX_INVOICE, APPROVAL_DOC], horizontal=True, key=f"kind_{key}")
    with st.form(f"ev_{key}_{kind}", clear_on_submit=True):
        vals = {}
        if kind == TAX_INVOICE:
            vals["승인번호"] = st.text_input("승인번호")
            vals["공급자사업자번호"] = st.text_input("공급자 사업자번호")
            vals["작성일"] = st.text_input("작성일 (YYYY-MM-DD)")
            money = {k: st.text_input(k) for k in ("공급가액", "세액", "합계")}
        else:
            vals["문서번호"] = st.text_input("문서번호")
            vals["승인상태"] = st.selectbox("승인상태", ["승인", "결재완료", "진행중", "반려", "미확인"])
            vals["목적"] = st.text_input("목적")
            money = {"금액": st.text_input("금액")}
            vals["기간시작"] = st.text_input("기간 시작 (YYYY-MM-DD)")
            vals["기간종료"] = st.text_input("기간 종료 (YYYY-MM-DD)")
        approval_id = st.text_input("결재문서 ID (있으면)")
        file_name = st.text_input("원본 파일명 (참고용)")
        if st.form_submit_button("수기 등록 후 이 전표에 연결"):
            errs = []
            for k, v in money.items():
                amount, stt = parse_amount(v)
                if stt == "오류":
                    errs.append(f"{k} 값을 해석할 수 없습니다: {v}")
                vals[k] = amount
            if errs:
                st.error(" / ".join(errs))
            else:
                vals["출처"] = "수기입력"
                doc = store.add_doc(kind, {k: v for k, v in vals.items() if v not in (None, "")}, user=state.reviewer,
                                    file_name=file_name, approval_doc_id=approval_id)
                store.link(key, doc.doc_id, BY_MANUAL, user=state.reviewer)
                state.rerun_checks()
                flash_and_rerun(f"{kind}를 등록·연결하고 다시 검사했습니다.")
