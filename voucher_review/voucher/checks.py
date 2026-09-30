"""전표 검사 규칙 엔진 (1차).

검사 영역: 데이터 점검 / 금액 검사 / 계정 검사 / 적요 검사 / 증빙 검사
상태: 통과 / 오류 / 확인 필요 / 미검사 / 해당 없음
요약: 오류 > 확인 필요 > 미검사 > 설정된 검사 통과  (검사 통과는 최종 회계 적정성 보증이 아님)

원칙: 설정된 기준이 없으면 비교 값을 만들지 않고 '미검사'로 남긴다.
"""

from __future__ import annotations

import datetime as dt
import re
from dataclasses import asdict, dataclass
from decimal import ROUND_DOWN, ROUND_HALF_UP, ROUND_UP, Decimal, InvalidOperation

import pandas as pd

from .accounts import SIDE_BOTH, SIDE_CREDIT, SIDE_DEBIT, SIDE_ERROR, SIDE_NONE, ZERO, with_side
from .columns import EXPECTED_COLUMNS
from .evidence import (
    APPROVAL_DOC,
    CHECK_FIELDS,
    LINK_CANDIDATE,
    LINK_CONFIRMED,
    LINK_FAILED,
    TAX_INVOICE,
    ManualEvidenceStore,
)
from .parser import parse_amount
from .settings import settings_version

PASS, ERROR, REVIEW, UNCHECKED, NA = "통과", "오류", "확인 필요", "미검사", "해당 없음"
STATUSES = [PASS, ERROR, REVIEW, UNCHECKED, NA]
SUMMARY_PASS = "설정된 검사 통과"
SUMMARY_ORDER = [ERROR, REVIEW, UNCHECKED, SUMMARY_PASS]

AREA_DATA, AREA_AMOUNT, AREA_ACCOUNT, AREA_MEMO, AREA_EVIDENCE = "데이터 점검", "금액 검사", "계정 검사", "적요 검사", "증빙 검사"
AREAS = [AREA_DATA, AREA_AMOUNT, AREA_ACCOUNT, AREA_MEMO, AREA_EVIDENCE]

_PRIORITY = {ERROR: 0, REVIEW: 1, UNCHECKED: 2, PASS: 3, NA: 4}
_ROUNDING = {"버림": ROUND_DOWN, "반올림": ROUND_HALF_UP, "올림": ROUND_UP}
APPROVED_STATES = {"승인", "결재완료", "완료", "승인완료"}
ERP_FLAG_NOTE = "ERP 첨부 표시 FALSE, 다우오피스 증빙 미확인"


@dataclass
class CheckResult:
    전표키: str
    검사영역: str
    검사코드: str
    상태: str
    실제값: str = ""
    비교값: str = ""  # 기준이 없으면 비워 둔다
    차이: str = ""
    근거: str = ""
    적용규칙버전: str = ""
    원본행번호: str = ""


def area_status(statuses: list[str]) -> str:
    if not statuses:
        return UNCHECKED
    return min(statuses, key=lambda s: _PRIORITY[s])


def summary_status(area_statuses: list[str]) -> str:
    """오류 → 확인 필요 → 미검사 → 설정된 검사 통과. 모든 영역이 필수 검사."""
    if ERROR in area_statuses:
        return ERROR
    if REVIEW in area_statuses:
        return REVIEW
    if UNCHECKED in area_statuses:
        return UNCHECKED
    return SUMMARY_PASS


def _s(v) -> str:
    if v is None:
        return ""
    if isinstance(v, Decimal):
        return f"{v:,.0f}" if v == v.to_integral_value() else f"{v:,}"
    return str(v)


def _rows(recs: list[dict]) -> str:
    return ",".join(str(r["_엑셀행"]) for r in recs)


def group_records(data: pd.DataFrame) -> dict[str, list[dict]]:
    """전표키 → 행 레코드 목록 (원본 순서 유지). DataFrame→dict 변환은 한 번만."""
    groups: dict[str, list[dict]] = {}
    for i, r in enumerate(data.to_dict("records")):
        r["_i"] = i
        groups.setdefault(r["전표키"], []).append(r)
    return groups


def _amt(v) -> Decimal:
    return v if isinstance(v, Decimal) else ZERO


def _norm(t) -> str:
    return re.sub(r"\s+", " ", str(t or "")).strip()


def _as_date(v) -> dt.date | None:
    if isinstance(v, dt.date):
        return v
    try:
        return dt.date.fromisoformat(str(v)) if v else None
    except ValueError:
        return None


# ------------------------------------------------------------------ 데이터 점검 ----
def check_data(data: pd.DataFrame, groups: dict[str, list[dict]], settings: dict, ver: str) -> list[CheckResult]:
    out: list[CheckResult] = []
    required = settings.get("required_fields", [])
    # 분개 식별자 중복: 회계단위+전표관리단위+전표기표번호 (삭제하지 않고 표시만)
    ident = data["회계단위"].astype(str) + "|" + data["전표관리단위"].astype(str) + "|" + data["전표기표번호"].astype(str)
    dup_ident = set(((ident.duplicated(keep=False)) & (data["전표기표번호"] != "")).to_numpy().nonzero()[0])
    # 완전 동일 행 (엑셀행 제외 전체 열 동일)
    raw_cols = [c for c in EXPECTED_COLUMNS if c in data.columns]
    dup_full = set(data[raw_cols].astype(str).duplicated(keep=False).to_numpy().nonzero()[0])

    for key, recs in groups.items():
        res = []
        for r in recs:
            row = str(r["_엑셀행"])
            missing = [c for c in required if r.get(c) in (None, "")]
            if missing:
                res.append(CheckResult(key, AREA_DATA, "D01", ERROR, ", ".join(missing), "", "",
                                       "필수 값 누락", ver, row))
            side = r["차대구분"]
            if side == SIDE_ERROR:
                bad = [c for c in ("차변금액", "대변금액") if r.get(f"{c}_상태") == "오류"]
                res.append(CheckResult(key, AREA_DATA, "D05", ERROR, ", ".join(bad), "", "",
                                       "금액을 해석할 수 없음 (0 으로 대체하지 않음)", ver, row))
            elif side == SIDE_BOTH:
                res.append(CheckResult(key, AREA_DATA, "D04", ERROR, f"차변 {_s(r['차변금액'])} / 대변 {_s(r['대변금액'])}",
                                       "", "", "한 행에 차변·대변 동시 입력", ver, row))
            elif side == SIDE_NONE:
                res.append(CheckResult(key, AREA_DATA, "D06", REVIEW, "차변·대변 모두 0 또는 빈값", "", "",
                                       "금액 없는 분개 행", ver, row))
        di = [r for r in recs if r["_i"] in dup_ident]
        if di:
            res.append(CheckResult(key, AREA_DATA, "D02", REVIEW, ", ".join(sorted({r["전표기표번호"] for r in di})),
                                   "", "", "동일 분개 식별자(전표기표번호) 중복 — 자동 삭제하지 않음", ver, _rows(di)))
        df_rows = [r for r in recs if r["_i"] in dup_full]
        if df_rows:
            res.append(CheckResult(key, AREA_DATA, "D03", REVIEW, f"{len(df_rows)}행", "", "",
                                   "모든 열이 같은 중복 행 — 추출 중복 여부 확인 (자동 삭제하지 않음)", ver, _rows(df_rows)))
        if not res:
            res.append(CheckResult(key, AREA_DATA, "D00", PASS, f"{len(recs)}행", "", "", "데이터 점검 이상 없음", ver, _rows(recs)))
        out.extend(res)
    return out


# ------------------------------------------------------------------ 금액 검사 ----
def is_vat_row(r: dict, vat: dict) -> bool:
    if r.get("계정과목코드") in set(vat.get("account_codes", [])):
        return True
    name = str(r.get("계정과목", ""))
    return any(k and k in name for k in vat.get("name_keywords", []))


def vat_row_amount(r: dict) -> Decimal:
    return _amt(r["차변금액"]) if r["차대구분"] == SIDE_DEBIT else _amt(r["대변금액"])


def check_vat_row(key: str, r: dict, vat: dict, ver: str) -> CheckResult:
    """부가세 행 '전표 내부 계산 점검' (실제 증빙 대조 아님)."""
    row = str(r["_엑셀행"])
    code = r.get("계정과목코드", "")
    evidence_val = _norm(r.get("증빙", ""))
    tax = vat_row_amount(r)

    def res(status, reason, actual="", compare="", diff=""):
        return CheckResult(key, AREA_AMOUNT, "V01", status, actual, compare, diff,
                           f"[전표 내부 계산 점검] {reason}", ver, row)

    if any(k and k in evidence_val for k in vat.get("excluded_evidence_keywords", [])):
        return res(NA, f"특수 증빙 유형('{evidence_val}') — 일반과세 규칙으로 판정하지 않음", _s(tax))
    if r["차대구분"] not in (SIDE_DEBIT, SIDE_CREDIT):
        return res(UNCHECKED, "부가세 행 금액 판정 불가 (0/빈값/차대동시/오류)", _s(tax))
    missing = []
    col = vat.get("supply_columns", {}).get(code)
    if not col:
        missing.append("공급가액 열 매핑")
    general = vat.get("general_evidence_values", [])
    if not general:
        missing.append("일반과세 증빙 유형")
    if not vat.get("rounding"):
        missing.append("반올림 기준")
    if missing:
        return res(UNCHECKED, "설정 필요: " + ", ".join(missing) + " (관리항목2는 공급가액 '후보'일 뿐 확인 전 사용 안 함)",
                   _s(tax))
    if evidence_val not in general:
        return res(UNCHECKED, f"증빙 유형 '{evidence_val or '(빈값)'}' 이 일반과세 대상으로 설정되지 않음", _s(tax))
    supply, st = parse_amount(r.get(col))
    if st == "빈값":
        return res(UNCHECKED, f"공급가액 열({col}) 값 없음", _s(tax))
    if st == "오류":
        return res(REVIEW, f"공급가액 열({col}) 값 해석 불가: {r.get(col)}", _s(tax))
    try:
        rate = Decimal(str(vat.get("rate", "0.1")))
        tol = Decimal(str(vat.get("tolerance", "0") or "0"))
    except InvalidOperation:
        return res(UNCHECKED, "세율/허용차이 설정 값 오류", _s(tax))
    expected = (supply * rate).quantize(Decimal("1"), rounding=_ROUNDING[vat["rounding"]])
    diff = tax - expected
    status = PASS if abs(diff) <= tol else REVIEW
    return res(status, f"공급가액({col}) {_s(supply)} × {rate} ({vat['rounding']}) · 허용차이 {_s(tol)}원",
               _s(tax), _s(expected), _s(diff))


def _totals(recs: list[dict]) -> tuple[Decimal, Decimal, bool]:
    ok = [r for r in recs if r["차대구분"] != SIDE_ERROR]
    d = sum((_amt(r["차변금액"]) for r in ok), ZERO)
    c = sum((_amt(r["대변금액"]) for r in ok), ZERO)
    return d, c, len(ok) == len(recs)


def check_amount(groups: dict[str, list[dict]], settings: dict, ver: str) -> list[CheckResult]:
    out = []
    vat = settings.get("vat", {})
    for key, recs in groups.items():
        d, c, all_ok = _totals(recs)
        if not all_ok:
            out.append(CheckResult(key, AREA_AMOUNT, "A01", UNCHECKED, f"차변 {_s(d)} / 대변 {_s(c)}", "", "",
                                   "해석 불가 금액 행이 있어 차대 합계를 확정할 수 없음", ver, _rows(recs)))
        elif d == c:
            out.append(CheckResult(key, AREA_AMOUNT, "A01", PASS, _s(d), _s(c), "0", "차변합계 = 대변합계 (원화)", ver, _rows(recs)))
        else:
            out.append(CheckResult(key, AREA_AMOUNT, "A01", ERROR, _s(d), _s(c), _s(d - c),
                                   "차대변 불일치 또는 추출 누락 확인 (필터된 불완전 전표일 수 있음)", ver, _rows(recs)))
        vat_rows = [r for r in recs if is_vat_row(r, vat)]
        if not vat_rows:
            out.append(CheckResult(key, AREA_AMOUNT, "V00", NA, "", "", "", "부가세 행 없음", ver, ""))
        for r in vat_rows:
            out.append(check_vat_row(key, r, vat, ver))
    return out


# ------------------------------------------------------------------ 계정 검사 ----
def _rule_applies(rule: dict, r: dict) -> bool:
    if not rule.get("활성"):
        return False
    kws = [k for k in rule.get("적요조건", []) if k]
    if kws and not any(_norm(k) in _norm(r.get("적요")) for k in kws):
        return False
    types = [t for t in rule.get("적용전표유형", []) if t]
    if types and r.get("전표분개유형") not in types:
        return False
    side = rule.get("차대구분", "전체") or "전체"
    if side != "전체" and r["차대구분"] != side:
        return False
    day = r.get("회계일")
    start, end = _as_date(rule.get("적용시작일")), _as_date(rule.get("적용종료일"))
    if day and start and day < start:
        return False
    if day and end and day > end:
        return False
    return True


def _rule_ver(rule: dict) -> str:
    return f"{rule.get('규칙ID', '')}@v{rule.get('버전', '')}"


def check_account(groups: dict[str, list[dict]], settings: dict, ver: str) -> list[CheckResult]:
    """회사 승인 규칙으로만 검사. 규칙이 없는 행은 미검사. 거래처로 계정을 확정하지 않는다."""
    out = []
    rules = [r for r in settings.get("account_rules", []) if r.get("활성")]
    for key, recs in groups.items():
        res, uncovered = [], []
        for r in recs:
            if r["차대구분"] not in (SIDE_DEBIT, SIDE_CREDIT):
                continue
            hits = [rule for rule in rules if _rule_applies(rule, r)]
            if not hits:
                uncovered.append(r)
                continue
            allowed_by = [h for h in hits if r["계정과목코드"] in set(h.get("허용계정코드", []))]
            vers = ", ".join(_rule_ver(h) for h in hits)
            if allowed_by:
                res.append(CheckResult(key, AREA_ACCOUNT, "C01", PASS, f"{r['계정과목코드']} {r['계정과목']}",
                                       ", ".join(allowed_by[0].get("허용계정코드", [])), "",
                                       allowed_by[0].get("근거", ""), f"{ver} / {vers}", str(r["_엑셀행"])))
            else:
                allowed = sorted({c for h in hits for c in h.get("허용계정코드", [])})
                res.append(CheckResult(key, AREA_ACCOUNT, "C01", REVIEW, f"{r['계정과목코드']} {r['계정과목']}",
                                       ", ".join(allowed), "", "회사 규칙의 허용 계정이 아님: " +
                                       " / ".join(h.get("근거", "") for h in hits), f"{ver} / {vers}", str(r["_엑셀행"])))
        # 필수 상대계정(조합)
        for rule in rules:
            need = set(rule.get("필수상대계정코드", []))
            matched = [r for r in recs if r["차대구분"] in (SIDE_DEBIT, SIDE_CREDIT) and _rule_applies(rule, r)]
            if not need or not matched:
                continue
            codes = {r["계정과목코드"] for r in recs}
            if codes & need:
                res.append(CheckResult(key, AREA_ACCOUNT, "C02", PASS, ", ".join(sorted(codes & need)), ", ".join(sorted(need)),
                                       "", f"필수 계정 조합 충족: {rule.get('근거', '')}", f"{ver} / {_rule_ver(rule)}",
                                       ",".join(str(r["_엑셀행"]) for r in matched)))
            else:
                res.append(CheckResult(key, AREA_ACCOUNT, "C02", REVIEW, ", ".join(sorted(codes)), ", ".join(sorted(need)),
                                       "", f"필수 계정 조합 없음: {rule.get('근거', '')}", f"{ver} / {_rule_ver(rule)}",
                                       ",".join(str(r["_엑셀행"]) for r in matched)))
        if uncovered:
            res.append(CheckResult(key, AREA_ACCOUNT, "C00", UNCHECKED, f"{len(uncovered)}행", "", "",
                                   "적용되는 회사 승인 계정 규칙 없음", ver, ",".join(str(r["_엑셀행"]) for r in uncovered)))
        if not res:
            res.append(CheckResult(key, AREA_ACCOUNT, "C00", UNCHECKED, "", "", "", "검사 가능한 분개 행 없음", ver, _rows(recs)))
        out.extend(res)
    return out


# ------------------------------------------------------------------ 적요 검사 ----
def check_memo(groups: dict[str, list[dict]], settings: dict, ver: str) -> list[CheckResult]:
    """적요 빈칸 + 사용자가 설정한 거래유형별 필수 적요 정보만 점검. 글자 수로 적정성을 단정하지 않음."""
    memo_cfg = settings.get("memo", {})
    blank_status = memo_cfg.get("blank_status", ERROR)
    rules = [r for r in memo_cfg.get("rules", []) if r.get("활성")]
    out = []
    for key, recs in groups.items():
        res = []
        for r in recs:
            memo = _norm(r.get("적요"))
            if not memo:
                res.append(CheckResult(key, AREA_MEMO, "M01", blank_status, "(빈칸)", "", "", "적요 빈칸", ver,
                                       str(r["_엑셀행"])))
                continue
            for rule in rules:
                codes = set(rule.get("대상계정코드", []))
                if codes and r["계정과목코드"] not in codes:
                    continue
                cond = [k for k in rule.get("적요조건", []) if k]
                if cond and not any(_norm(k) in memo for k in cond):
                    continue
                need = [k for k in rule.get("필수포함", []) if k]
                lack = [k for k in need if _norm(k) not in memo]
                status = REVIEW if lack else PASS
                res.append(CheckResult(key, AREA_MEMO, "M02", status, memo, ", ".join(need),
                                       ", ".join(lack) and f"누락: {', '.join(lack)}", rule.get("근거", ""),
                                       f"{ver} / {_rule_ver(rule)}", str(r["_엑셀행"])))
        if not res:
            res.append(CheckResult(key, AREA_MEMO, "M00", PASS, "", "", "",
                                   "적요 빈칸 없음" + ("" if rules else " (거래유형별 필수 적요 기준 미설정 — 적정성은 판단하지 않음)"),
                                   ver, _rows(recs)))
        out.extend(res)
    return out


# ------------------------------------------------------------------ 증빙 검사 ----
def check_evidence(groups: dict[str, list[dict]], settings: dict, ver: str,
                   store: ManualEvidenceStore | None) -> list[CheckResult]:
    """증빙은 연결·확인된 문서로만 검사. ERP 파일첨부여부·'증빙' 열 텍스트는 증거로 쓰지 않음."""
    ev_cfg = settings.get("evidence", {})
    vat = settings.get("vat", {})
    store = store or ManualEvidenceStore()
    out = []
    for key, recs in groups.items():
        flags = {r.get("파일첨부여부_구분") for r in recs}
        flag_note = ERP_FLAG_NOTE if "문자열 FALSE" in flags or "불리언 FALSE" in flags else "다우오피스 증빙 미확인"
        vat_rows = [r for r in recs if is_vat_row(r, vat)]
        required = []
        if ev_cfg.get("require_tax_invoice_for_vat", True) and vat_rows:
            required.append(TAX_INVOICE)
        if ev_cfg.get("require_approval_doc", True):
            required.append(APPROVAL_DOC)
        if not required:
            out.append(CheckResult(key, AREA_EVIDENCE, "E00", NA, "", "", "", "설정상 요구 증빙 없음", ver, ""))
            continue
        links = store.links_for(key)
        for doc_type in required:
            typed = [(l, store.get_doc(l.doc_id)) for l in links]
            typed = [(l, d) for l, d in typed if d is None or d.doc_type == doc_type]
            confirmed = [d for l, d in typed if l.status == LINK_CONFIRMED and d is not None]
            code = "E01" if doc_type == TAX_INVOICE else "E02"
            if not confirmed:
                if any(l.status == LINK_CANDIDATE for l, _ in typed):
                    reason = f"{doc_type} 후보만 있음 — 수동 연결 확정 필요"
                elif any(l.status == LINK_FAILED or d is None for l, d in typed):
                    reason = f"{doc_type} 연결 실패"
                else:
                    reason = f"{doc_type} 미제공/미연결 · {flag_note}"
                out.append(CheckResult(key, AREA_EVIDENCE, code, UNCHECKED, "", "", "", reason, ver, ""))
                continue
            failed = [d for d in confirmed if d.recognition_failed]
            pending = [(d, d.unusable_fields(CHECK_FIELDS[doc_type])) for d in confirmed if not d.recognition_failed]
            usable = [d for d, miss in pending if not miss]
            if not usable:
                why = "인식 실패" if failed else "필수 값 수동 확인 전 또는 누락: " + ", ".join(
                    sorted({m for _, miss in pending for m in miss}))
                out.append(CheckResult(key, AREA_EVIDENCE, code, UNCHECKED, "", "", "", f"{doc_type} {why}", ver, ""))
                continue
            if doc_type == TAX_INVOICE:
                out.extend(_compare_tax_invoice(key, usable, vat_rows, vat, ver))
            else:
                for d in usable:
                    state = _norm(d.value("승인상태"))
                    status = PASS if state in APPROVED_STATES else REVIEW
                    out.append(CheckResult(key, AREA_EVIDENCE, code, status, f"{d.value('문서번호')} / {state}",
                                           "승인 상태", "", f"[실제 증빙 대조] 품의서 승인 상태 확인 (목적: {d.value('목적') or '-'})",
                                           ver, ""))
    return out


def _compare_tax_invoice(key, docs, vat_rows, vat, ver) -> list[CheckResult]:
    """[실제 증빙 대조] 연결·확인된 세금계산서 세액 합계 vs 전표 부가세 행 합계."""
    inv_tax = sum((Decimal(str(d.value("세액"))) for d in docs), ZERO)
    rows = ",".join(str(r["_엑셀행"]) for r in vat_rows)
    ids = ", ".join(str(d.value("승인번호")) for d in docs)
    out = []
    if not vat_rows:
        out.append(CheckResult(key, AREA_EVIDENCE, "E01", REVIEW, _s(inv_tax), "", "",
                               f"[실제 증빙 대조] 세금계산서({ids})는 연결됐으나 전표에 부가세 행이 없음", ver, ""))
        return out
    v_tax = sum((vat_row_amount(r) for r in vat_rows), ZERO)
    diff = v_tax - inv_tax
    out.append(CheckResult(key, AREA_EVIDENCE, "E01", PASS if diff == 0 else REVIEW, _s(v_tax), _s(inv_tax), _s(diff),
                           f"[실제 증빙 대조] 전표 부가세 합계 vs 세금계산서 세액 합계 ({ids})", ver, rows))
    cols = {vat.get("supply_columns", {}).get(r["계정과목코드"]) for r in vat_rows}
    if None in cols or not cols:
        out.append(CheckResult(key, AREA_EVIDENCE, "E03", UNCHECKED, "", "", "",
                               "[실제 증빙 대조] 공급가액 열 매핑 미설정 — 공급가액 대조 안 함", ver, rows))
    else:
        supplies = [parse_amount(r.get(vat["supply_columns"][r["계정과목코드"]])) for r in vat_rows]
        if any(st != "정상" for _, st in supplies):
            out.append(CheckResult(key, AREA_EVIDENCE, "E03", UNCHECKED, "", "", "",
                                   "[실제 증빙 대조] 전표 공급가액 값 없음/해석 불가", ver, rows))
        else:
            v_sup = sum((a for a, _ in supplies), ZERO)
            i_sup = sum((Decimal(str(d.value("공급가액"))) for d in docs), ZERO)
            out.append(CheckResult(key, AREA_EVIDENCE, "E03", PASS if v_sup == i_sup else REVIEW, _s(v_sup), _s(i_sup),
                                   _s(v_sup - i_sup), "[실제 증빙 대조] 전표 공급가액 vs 세금계산서 공급가액", ver, rows))
    return out


# ------------------------------------------------------------------ 실행 ----
def run_checks(df: pd.DataFrame, settings: dict, store: ManualEvidenceStore | None = None) -> pd.DataFrame:
    data = with_side(df)
    groups = group_records(data)
    ver = settings_version(settings)
    results = (check_data(data, groups, settings, ver) + check_amount(groups, settings, ver)
               + check_account(groups, settings, ver) + check_memo(groups, settings, ver)
               + check_evidence(groups, settings, ver, store))
    return pd.DataFrame([asdict(r) for r in results], columns=list(CheckResult.__dataclass_fields__))


def _first(recs: list[dict], col: str):
    return next((r.get(col) for r in recs if r.get(col) not in (None, "")), "")


def voucher_summary(df: pd.DataFrame, results: pd.DataFrame) -> pd.DataFrame:
    """전표별 1행. 대표값(기표자·부서 등)은 전표에서 처음 나오는 값 기준."""
    groups = group_records(with_side(df))
    by_key: dict[str, list[dict]] = {}
    for r in results.to_dict("records"):
        by_key.setdefault(r["전표키"], []).append(r)
    rows = []
    for key, recs in groups.items():
        d, c, _ = _totals(recs)
        res = by_key.get(key, [])
        areas = {a: area_status([r["상태"] for r in res if r["검사영역"] == a]) for a in AREAS}
        reasons = sorted((r for r in res if r["상태"] in (ERROR, REVIEW, UNCHECKED)), key=lambda r: _PRIORITY[r["상태"]])
        memos = [r["적요"] for r in recs if r.get("적요")]
        rows.append({
            "전표키": key, "회계단위": _first(recs, "회계단위"), "전표관리단위": _first(recs, "전표관리단위"),
            "기표번호": _first(recs, "기표번호"), "회계일": _first(recs, "회계일") or None, "기표자": _first(recs, "기표자"),
            "기표부서": _first(recs, "기표부서"), "결재상태": _first(recs, "전자결재진행상태"),
            "전표유형": _first(recs, "전표분개유형"),
            "대표적요": max(dict.fromkeys(memos), key=memos.count) if memos else "", "행수": len(recs),
            "차변합계": d, "대변합계": c, "차이": d - c, **areas,
            "검사결과": summary_status(list(areas.values())),
            "사유": " / ".join(dict.fromkeys(f"[{r['검사영역']}] {r['근거']}" for r in reasons))[:500],
        })
    return pd.DataFrame(rows)


def summary_counts(summary: pd.DataFrame) -> dict[str, int]:
    """네 상태는 중복 없이 전체 전표 수와 일치."""
    counts = {s: int((summary["검사결과"] == s).sum()) for s in SUMMARY_ORDER}
    counts["전체"] = len(summary)
    assert sum(counts[s] for s in SUMMARY_ORDER) == counts["전체"]
    return counts


def sort_error_first(summary: pd.DataFrame) -> pd.DataFrame:
    order = {s: i for i, s in enumerate(SUMMARY_ORDER)}
    return summary.assign(_o=summary["검사결과"].map(order)).sort_values(["_o", "회계일", "기표번호"], na_position="last").drop(columns="_o")
