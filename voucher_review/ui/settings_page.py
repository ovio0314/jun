"""검사 설정 화면. 설정되지 않은 기준으로는 검사하지 않는다(미검사)."""

from __future__ import annotations

import copy
import json

import pandas as pd
import streamlit as st

from voucher.checks import is_vat_row
from voucher.columns import MANAGEMENT_ITEM_COLUMNS
from voucher.qa import load_guide, validate_guide
from voucher.settings import (
    ACCOUNT_RULE_FIELDS,
    MEMO_RULE_FIELDS,
    ROUNDING_METHODS,
    default_settings,
    example_rules,
    save_settings,
    settings_version,
)

from .common import flash_and_rerun, join_list, split_list

LIST_FIELDS_ACCOUNT = ["적요조건", "적용전표유형", "허용계정코드", "필수상대계정코드"]
LIST_FIELDS_MEMO = ["대상계정코드", "적요조건", "필수포함"]


def rules_to_frame(rules: list[dict], fields: list[str], list_fields: list[str]) -> pd.DataFrame:
    rows = []
    for r in rules:
        row = {f: r.get(f, "") for f in fields}
        for f in list_fields:
            row[f] = join_list(r.get(f, []))
        row["활성"] = bool(r.get("활성", False))
        row["예시"] = bool(r.get("예시", False))
        rows.append(row)
    return pd.DataFrame(rows, columns=fields).astype({f: object for f in fields if f not in ("활성", "예시")})


def frame_to_rules(df: pd.DataFrame, list_fields: list[str]) -> list[dict]:
    out = []
    for r in df.to_dict("records"):
        if not str(r.get("규칙ID") or "").strip():
            continue
        rule = {k: ("" if v is None or (isinstance(v, float) and v != v) else v) for k, v in r.items()}
        for f in list_fields:
            rule[f] = split_list(rule.get(f))
        rule["활성"] = bool(rule.get("활성"))
        rule["예시"] = bool(rule.get("예시"))
        rule["버전"] = str(rule.get("버전") or "1")
        if "차대구분" in rule and rule["차대구분"] not in ("차변", "대변", "전체"):
            rule["차대구분"] = "전체"
        out.append(rule)
    return out


def render_settings(state, settings_path, local_guide_path) -> None:
    s = copy.deepcopy(state.settings)
    st.markdown(f"**현재 설정 버전:** `{settings_version(state.settings)}` — 결과의 '적용규칙버전'에 기록됩니다.")
    st.caption("보수적 기본값: 회사가 기준을 설정하기 전에는 계정 검사·부가세 계산 점검·거래유형별 적요 검사를 '미검사'로 둡니다.")
    df = state.result.df if state.result is not None and not state.result.errors else None

    # ---------------- 부가세
    st.subheader("부가세 (전표 내부 계산 점검)")
    vat = s["vat"]
    codes_in_data = sorted(df["계정과목코드"].unique()) if df is not None else []
    names = dict(zip(df["계정과목코드"], df["계정과목"])) if df is not None else {}
    vat["account_codes"] = st.multiselect(
        "부가세 계정코드", sorted(set(codes_in_data) | set(vat["account_codes"])), default=vat["account_codes"], placeholder="선택 (없으면 계정명 키워드로만 식별)",
        format_func=lambda c: f"{c} {names.get(c, '')}")
    vat["name_keywords"] = split_list(st.text_input("부가세 계정명 키워드 (쉼표 구분)", join_list(vat["name_keywords"])))
    if df is not None:
        vat_codes = sorted({r["계정과목코드"] for r in df.to_dict("records") if is_vat_row(r, vat)})
        st.markdown("**공급가액 열 매핑** — 계정별로 관리항목 의미가 다릅니다. 확인한 열만 지정하세요.")
        for code in vat_codes:
            sub = df[df["계정과목코드"] == code]
            samples = {c: [v for v in sub[c] if v][:3] for c in MANAGEMENT_ITEM_COLUMNS}
            hint = "관리항목2는 공급가액 **후보**일 뿐 — 값 예시를 보고 확인 후 지정" if samples.get("관리항목2") else ""
            opts = ["(미설정)"] + MANAGEMENT_ITEM_COLUMNS
            cur = vat["supply_columns"].get(code, "(미설정)")
            pick = st.selectbox(f"{code} {names.get(code, '')} 공급가액 열", opts, index=opts.index(cur) if cur in opts else 0,
                                key=f"supply_{code}", help=hint)
            st.caption("값 예시: " + (" · ".join(f"{c}={', '.join(map(str, v))}" for c, v in samples.items() if v) or "(관리항목 값 없음)") + (f"  ({hint})" if hint else ""))
            if pick == "(미설정)":
                vat["supply_columns"].pop(code, None)
            else:
                vat["supply_columns"][code] = pick
        ev_values = sorted({v for v in df["증빙"] if v})
    else:
        st.caption("전표 파일을 불러오면 부가세 계정별 공급가액 열을 매핑할 수 있습니다.")
        ev_values = []
    vat["general_evidence_values"] = st.multiselect(
        "일반과세 내부 계산 대상 '증빙' 값", sorted(set(ev_values) | set(vat["general_evidence_values"])),
        default=vat["general_evidence_values"], placeholder="선택 안 함 → 부가세 계산 점검 미검사",
        help="'증빙' 열 텍스트는 유형 구분에만 쓰며 실제 첨부 증거로 쓰지 않습니다.")
    vat["excluded_evidence_keywords"] = split_list(st.text_input(
        "특수 증빙 제외 키워드 (불공제·영세율·면세·수정·분할 등, 쉼표 구분)", join_list(vat["excluded_evidence_keywords"])))
    c = st.columns(3)
    rounding_opts = [""] + list(ROUNDING_METHODS)
    vat["rounding"] = c[0].selectbox("반올림 기준", rounding_opts, index=rounding_opts.index(vat["rounding"]),
                                     format_func=lambda x: x or "(미설정 → 미검사)")
    vat["rate"] = c[1].text_input("세율", vat["rate"])
    vat["tolerance"] = c[2].text_input("허용 차이(원)", vat["tolerance"])

    # ---------------- 계정 규칙
    st.subheader("회사 승인 계정 규칙")
    st.caption("규칙이 적용되지 않는 분개는 계정 검사 '미검사'. 거래처만으로 계정을 확정하지 않습니다. 목록 값은 쉼표로 구분. "
               "차대구분: 차변/대변/전체. 날짜: YYYY-MM-DD.")
    if st.button("합성 데이터용 예시 규칙 추가 (비활성 상태로)"):
        ex = example_rules()
        have = {r["규칙ID"] for r in s["account_rules"]}
        s["account_rules"] += [r for r in ex["account_rules"] if r["규칙ID"] not in have]
        have_m = {r["규칙ID"] for r in s["memo"]["rules"]}
        s["memo"]["rules"] += [r for r in ex["memo_rules"] if r["규칙ID"] not in have_m]
        state.settings = s
        st.rerun()
    acc = st.data_editor(rules_to_frame(s["account_rules"], ACCOUNT_RULE_FIELDS, LIST_FIELDS_ACCOUNT), num_rows="dynamic",
                         use_container_width=True, key="acc_rules",
                         column_config={"활성": st.column_config.CheckboxColumn(), "예시": st.column_config.CheckboxColumn(),
                                        "차대구분": st.column_config.SelectboxColumn(options=["전체", "차변", "대변"])})
    s["account_rules"] = frame_to_rules(acc, LIST_FIELDS_ACCOUNT)

    # ---------------- 적요
    st.subheader("적요 검사")
    opts = ["오류", "확인 필요"]
    s["memo"]["blank_status"] = st.selectbox("적요 빈칸 판정", opts, index=opts.index(s["memo"]["blank_status"]))
    st.caption("거래유형별 필수 적요 정보: 대상계정코드/적요조건에 해당하는 행의 적요에 '필수포함' 단어가 모두 있어야 합니다. "
               "글자 수만으로 적정성을 판단하지 않습니다.")
    memo = st.data_editor(rules_to_frame(s["memo"]["rules"], MEMO_RULE_FIELDS, LIST_FIELDS_MEMO), num_rows="dynamic",
                          use_container_width=True, key="memo_rules",
                          column_config={"활성": st.column_config.CheckboxColumn(), "예시": st.column_config.CheckboxColumn()})
    s["memo"]["rules"] = frame_to_rules(memo, LIST_FIELDS_MEMO)

    # ---------------- 증빙
    st.subheader("증빙 요구 기준")
    s["evidence"]["require_tax_invoice_for_vat"] = st.checkbox("부가세 행이 있는 전표는 세금계산서 필요",
                                                               s["evidence"]["require_tax_invoice_for_vat"])
    s["evidence"]["require_approval_doc"] = st.checkbox("모든 전표에 품의서 필요", s["evidence"]["require_approval_doc"])
    st.caption("증빙이 연결·확인되지 않으면 증빙 검사는 '미검사'로 남고, 전체 요약도 '설정된 검사 통과'가 되지 않습니다.")

    c = st.columns(3)
    if c[0].button("설정 적용 (이번 세션)", type="primary"):
        state.settings = s
        state.rerun_checks()
        flash_and_rerun(f"적용했습니다. 설정 버전 {settings_version(s)} 로 다시 검사했습니다.")
    if c[1].button("설정 적용 + 이 PC에 저장"):
        state.settings = s
        save_settings(s, settings_path)
        state.rerun_checks()
        flash_and_rerun(f"저장했습니다: {settings_path.name} (설정 버전 {settings_version(s)})")
    if c[2].button("기본값으로 되돌리기 (세션)"):
        state.settings = default_settings()
        flash_and_rerun("기본 설정으로 되돌렸습니다. 검사 실행을 다시 눌러 주세요.")

    # ---------------- 계정 질문 가이드
    st.divider()
    st.subheader("계정 질문 가이드 (제조업 참고 분개)")
    guide = st.session_state.get("guide") or load_guide(local_guide_path if local_guide_path.exists() else None)
    st.caption(f"버전 {guide.get('version', '-')} · {guide.get('기준', '')}")
    st.download_button("가이드 JSON 다운로드", json.dumps(guide, ensure_ascii=False, indent=2).encode("utf-8"),
                       file_name="account_guide.json", mime="application/json")
    up = st.file_uploader("회사 기준으로 수정한 가이드 JSON 올리기", type=["json"], key="guide_upload")
    if up is not None:
        try:
            new_guide = json.loads(up.getvalue().decode("utf-8"))
            validate_guide(new_guide)
        except (ValueError, UnicodeDecodeError) as e:
            st.error(f"가이드 JSON 을 읽을 수 없습니다: {e}")
        else:
            if st.button("가이드 적용 + 이 PC에 저장"):
                local_guide_path.parent.mkdir(exist_ok=True)
                local_guide_path.write_text(json.dumps(new_guide, ensure_ascii=False, indent=2), encoding="utf-8")
                st.session_state.guide = new_guide
                flash_and_rerun("회사 가이드를 저장·적용했습니다.")
    if local_guide_path.exists():
        ok = st.checkbox("저장된 회사 가이드 파일을 삭제하고 기본 가이드로 되돌리는 것에 동의합니다.")
        if st.button("저장된 가이드 삭제", disabled=not ok):
            local_guide_path.unlink()
            st.session_state.guide = load_guide()
            flash_and_rerun("저장된 가이드를 삭제하고 기본 가이드로 되돌렸습니다.")
