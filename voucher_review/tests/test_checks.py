"""필수 테스트: 규칙 엔진·상태 모델."""

import copy
from decimal import Decimal

import pytest

from make_synthetic import build_rows, write_xlsx
from voucher.checks import (
    AREA_ACCOUNT,
    AREA_AMOUNT,
    AREA_DATA,
    AREA_EVIDENCE,
    ERROR,
    NA,
    PASS,
    REVIEW,
    SUMMARY_ORDER,
    SUMMARY_PASS,
    UNCHECKED,
    run_checks,
    summary_counts,
    voucher_summary,
)
from voucher.evidence import APPROVAL_DOC, BY_MANUAL, TAX_INVOICE, ManualEvidenceStore
from voucher.parser import find_header_candidates, load_vouchers
from voucher.settings import default_settings, example_rules


def load(tmp_path, rows, name="t.xlsx"):
    p = write_xlsx(tmp_path / name, rows=rows)
    return load_vouchers(p, find_header_candidates(p)[0]).df


def run(df, settings=None, store=None):
    res = run_checks(df, settings or default_settings(), store)
    return res, voucher_summary(df, res)


def codes(res, key=None, area=None):
    r = res
    if key:
        r = r[r["전표키"] == key]
    if area:
        r = r[r["검사영역"] == area]
    return list(zip(r["검사코드"], r["상태"]))


def vat_settings(rounding="버림"):
    s = default_settings()
    s["vat"].update(account_codes=["13500", "25500"], supply_columns={"13500": "관리항목2"},
                    general_evidence_values=["세금계산서"], rounding=rounding)
    return s


# 원재료 매입 1건: 원재료 / 부가세대급금 / 외상매입금
def raw_purchase(supply=1_000_000, tax=None, evidence="세금계산서", no="V1", unit="본사"):
    rows = build_rows(n_vouchers=1)
    tax = supply // 10 if tax is None else tax
    rows[0].update(차변금액=supply, 관리항목2=str(supply))
    rows[1].update(차변금액=tax, 관리항목2=str(supply), 증빙=evidence)
    rows[2].update(대변금액=supply + tax)
    for r in rows:
        r.update(기표번호=no, 회계단위=unit)
    for i, r in enumerate(rows, 1):
        r["전표기표번호"] = f"{unit}-{no}-{i}"
    return rows


def test_balance_match_and_mismatch(tmp_path):
    ok = raw_purchase(no="OK")
    bad = raw_purchase(no="BAD")
    bad[2]["대변금액"] -= 1
    res, summ = run(load(tmp_path, ok + bad))
    s = summ.set_index("기표번호")
    assert s.loc["OK", "금액 검사"] in (PASS, UNCHECKED, NA)
    assert codes(res, "본사|회계|OK", AREA_AMOUNT)[0] == ("A01", PASS)
    assert codes(res, "본사|회계|BAD", AREA_AMOUNT)[0] == ("A01", ERROR)
    assert s.loc["BAD", "차이"] == Decimal(1)
    assert "차대변 불일치 또는 추출 누락 확인" in s.loc["BAD", "사유"]


def test_same_number_in_multiple_accounting_units(tmp_path):
    a = raw_purchase(no="X1", unit="본사")
    b = raw_purchase(no="X1", unit="공장")
    b[0]["차변금액"] += 5  # 공장 전표만 불일치
    res, summ = run(load(tmp_path, a + b))
    assert len(summ) == 2
    assert codes(res, "본사|회계|X1", AREA_AMOUNT)[0][1] == PASS
    assert codes(res, "공장|회계|X1", AREA_AMOUNT)[0][1] == ERROR


def test_blank_voucher_number_not_grouped(tmp_path):
    rows = raw_purchase(no="")
    df = load(tmp_path, rows)
    assert df["전표키"].nunique() == 3  # 행마다 따로
    res, summ = run(df)
    assert (summ["데이터 점검"] == ERROR).all()
    assert ("D01", ERROR) in codes(res, area=AREA_DATA)


def test_duplicate_rows_flagged_not_deleted(tmp_path):
    rows = raw_purchase(no="DUP")
    rows.append(copy.deepcopy(rows[0]))
    df = load(tmp_path, rows)
    assert len(df) == 4
    res, _ = run(df)
    data = codes(res, area=AREA_DATA)
    assert ("D02", REVIEW) in data and ("D03", REVIEW) in data
    assert codes(res, area=AREA_AMOUNT)[0] == ("A01", ERROR)  # 중복으로 차변 과다 → 불일치


def test_both_sides_one_row_is_error(tmp_path):
    rows = raw_purchase(no="BOTH")
    rows[0]["대변금액"] = 10
    rows[2]["대변금액"] += 0
    res, _ = run(load(tmp_path, rows))
    assert ("D04", ERROR) in codes(res, area=AREA_DATA)


def test_comma_numbers(tmp_path):
    rows = raw_purchase(supply=1_234_000, no="CM")
    for r in rows:
        for c in ("차변금액", "대변금액"):
            r[c] = f"{r[c]:,}" if r[c] else ""
    res, summ = run(load(tmp_path, rows))
    assert codes(res, area=AREA_AMOUNT)[0] == ("A01", PASS)
    assert summ["차변합계"].iloc[0] == Decimal("1357400")


def test_invalid_number_not_zero(tmp_path):
    rows = raw_purchase(no="BADNUM")
    rows[0]["차변금액"] = "1,000,O00"  # 영문 O
    res, summ = run(load(tmp_path, rows))
    assert ("D05", ERROR) in codes(res, area=AREA_DATA)
    assert codes(res, area=AREA_AMOUNT)[0] == ("A01", UNCHECKED)
    assert summ["검사결과"].iloc[0] == ERROR


def test_negative_amounts_not_error(tmp_path):
    rows = raw_purchase(no="NEG")
    for r in rows:
        r["차변금액"], r["대변금액"] = -r["차변금액"], -r["대변금액"]
    res, summ = run(load(tmp_path, rows))
    assert summ["데이터 점검"].iloc[0] == PASS
    assert codes(res, area=AREA_AMOUNT)[0] == ("A01", PASS)


def test_multiple_vat_rows_each_checked(tmp_path):
    rows = raw_purchase(no="MV", supply=1_000_000)
    extra = copy.deepcopy(rows[1])
    extra.update(차변금액=50_000, 관리항목2="500000", 전표기표번호="본사-MV-4", 행번호=4)
    rows.insert(2, extra)
    rows[-1]["대변금액"] += 50_000
    res, _ = run(load(tmp_path, rows), vat_settings())
    v01 = [s for c, s in codes(res, area=AREA_AMOUNT) if c == "V01"]
    assert v01 == [PASS, PASS]


@pytest.mark.parametrize("rounding, expected_status", [("버림", PASS), ("반올림", REVIEW), ("올림", REVIEW)])
def test_general_vat_rounding(tmp_path, rounding, expected_status):
    # 공급가액 12,345 × 0.1 = 1,234.5 → 버림 1,234 / 반올림·올림 1,235
    res, _ = run(load(tmp_path, raw_purchase(supply=12_345, tax=1_234, no="RD")), vat_settings(rounding))
    v01 = res[res["검사코드"] == "V01"].iloc[0]
    assert v01["상태"] == expected_status
    assert v01["실제값"] == "1,234"
    assert "전표 내부 계산 점검" in v01["근거"]


@pytest.mark.parametrize("evidence", ["세금계산서(불공제)", "영세율세금계산서", "수정세금계산서", "계산서(면세)"])
def test_special_evidence_excluded(tmp_path, evidence):
    res, _ = run(load(tmp_path, raw_purchase(supply=1000, tax=0, no="SP", evidence=evidence)), vat_settings())
    assert res[res["검사코드"] == "V01"]["상태"].iloc[0] == NA


def test_supply_mapping_not_configured_is_unchecked(tmp_path):
    s = vat_settings()
    s["vat"]["supply_columns"] = {}
    res, _ = run(load(tmp_path, raw_purchase(supply=1000, tax=999, no="NM")), s)
    v01 = res[res["검사코드"] == "V01"].iloc[0]
    assert v01["상태"] == UNCHECKED and v01["비교값"] == ""  # 기준 없는 비교값 없음
    assert "공급가액 열 매핑" in v01["근거"]


def test_default_settings_never_run_vat_calc(tmp_path):
    res, _ = run(load(tmp_path, raw_purchase(supply=1000, tax=999, no="DF")))
    assert res[res["검사코드"] == "V01"]["상태"].tolist() == [UNCHECKED]


def test_string_false_is_not_missing_evidence(loaded):
    res, summ = run(loaded.df)
    ev = res[res["검사영역"] == AREA_EVIDENCE]
    assert set(ev["상태"]) == {UNCHECKED}
    assert ev["근거"].str.contains("ERP 첨부 표시 FALSE, 다우오피스 증빙 미확인").all()
    assert not (summ["증빙 검사"] == ERROR).any()


def test_no_company_rules_means_unchecked(loaded):
    s = default_settings()
    s["account_rules"] = example_rules()["account_rules"]  # 예시 규칙은 비활성
    res, summ = run(loaded.df, s)
    assert set(summ["계정 검사"]) == {UNCHECKED}
    assert SUMMARY_PASS not in set(summ["검사결과"])


def test_active_rules_pass_and_flag(tmp_path):
    rows = raw_purchase(no="R1")
    wrong = raw_purchase(no="R2")
    wrong[2].update(계정과목코드="25300", 계정과목="미지급금")  # 원재료 매입인데 미지급금
    s = default_settings()
    s["account_rules"] = [dict(r, 활성=True) for r in example_rules()["account_rules"]]
    res, summ = run(load(tmp_path, rows + wrong), s)
    by = summ.set_index("기표번호")
    assert by.loc["R1", "계정 검사"] == PASS
    assert by.loc["R2", "계정 검사"] == REVIEW
    flagged = res[(res["전표키"] == "본사|회계|R2") & (res["상태"] == REVIEW)].iloc[0]
    assert "EX-A03@v1" in flagged["적용규칙버전"] and flagged["비교값"]


def test_evidence_unchecked_stays_in_summary(tmp_path):
    rows = raw_purchase(no="EV")
    s = vat_settings()
    s["account_rules"] = [dict(r, 활성=True) for r in example_rules()["account_rules"]]
    df = load(tmp_path, rows)
    res, summ = run(df, s)
    row = summ.iloc[0]
    assert [row[a] for a in ("데이터 점검", "금액 검사", "계정 검사", "적요 검사")] == [PASS] * 4
    assert row["증빙 검사"] == UNCHECKED and row["검사결과"] == UNCHECKED

    store = ManualEvidenceStore()
    key = df["전표키"].iloc[0]
    inv = store.add_doc(TAX_INVOICE, {"승인번호": "2026-001", "공급가액": Decimal(1_000_000), "세액": Decimal(100_000),
                                      "합계": Decimal(1_100_000)}, file_name="세금계산서.pdf")
    store.link(key, inv.doc_id, BY_MANUAL)
    res, summ = run(df, s, store)
    assert summ.iloc[0]["증빙 검사"] == UNCHECKED  # 품의서가 아직 없음
    doc = store.add_doc(APPROVAL_DOC, {"문서번호": "품의-1", "승인상태": "승인"})
    store.link(key, doc.doc_id, BY_MANUAL)
    res, summ = run(df, s, store)
    assert summ.iloc[0]["증빙 검사"] == PASS and summ.iloc[0]["검사결과"] == SUMMARY_PASS
    assert ("E01", PASS) in codes(res, area=AREA_EVIDENCE) and ("E03", PASS) in codes(res, area=AREA_EVIDENCE)


def test_summary_counts_no_double_count(loaded):
    res, summ = run(loaded.df)
    counts = summary_counts(summ)
    assert sum(counts[s] for s in SUMMARY_ORDER) == counts["전체"] == loaded.df["전표키"].nunique()
    assert summ["전표키"].is_unique
    assert sum(summ["차변합계"], Decimal(0)) == sum((v for v in loaded.df["차변금액"] if v), Decimal(0))


def test_checks_do_not_change_review_state(loaded):
    from voucher.review import content_hashes, init_reviews

    reviews = init_reviews(content_hashes(loaded.df))
    run(loaded.df)
    assert {r.검토상태 for r in reviews.values()} == {"미검토"}
