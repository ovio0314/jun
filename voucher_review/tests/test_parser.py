import datetime as dt
import hashlib
from decimal import Decimal

import pytest
from openpyxl import Workbook

from make_synthetic import build_rows, write_xlsx
from voucher.accounts import SIDE_ERROR, with_side
from voucher.columns import EXPECTED_COLUMNS
from voucher.parser import (
    VoucherFileError,
    classify_flag,
    find_header_candidates,
    load_vouchers,
    parse_amount,
)


@pytest.mark.parametrize(
    "value, expected, status",
    [
        ("1,234,000", Decimal("1234000"), "정상"),
        (" 1 234 ", Decimal("1234"), "정상"),
        ("12,000원", Decimal("12000"), "정상"),
        ("(1,000)", Decimal("-1000"), "정상"),
        ("△500", Decimal("-500"), "정상"),
        ("-500", Decimal("-500"), "정상"),
        (1000, Decimal("1000"), "정상"),
        (0.1, Decimal("0.1"), "정상"),
        (None, None, "빈값"),
        ("   ", None, "빈값"),
        ("abc", None, "오류"),
        ("1.2.3", None, "오류"),
        (True, None, "오류"),
    ],
)
def test_parse_amount(value, expected, status):
    amount, st = parse_amount(value)
    assert st == status
    assert amount == expected
    if amount is not None:
        assert isinstance(amount, Decimal)


def test_invalid_amount_is_not_zero():
    amount, status = parse_amount("확인필요")
    assert amount is None and status == "오류"


@pytest.mark.parametrize(
    "value, expected",
    [("FALSE", "문자열 FALSE"), ("true", "문자열 TRUE"), (True, "불리언 TRUE"), (False, "불리언 FALSE"),
     (None, "빈값"), ("", "빈값"), ("Y", "기타값(Y)")],
)
def test_classify_flag(value, expected):
    assert classify_flag(value) == expected


def test_header_detection_title_and_row(synthetic_path):
    cands = find_header_candidates(synthetic_path)
    assert len(cands) == 1
    c = cands[0]
    assert c.sheet == "Sheet1" and c.header_row == 2 and c.matched == len(EXPECTED_COLUMNS)
    assert c.title == "전표조건검색"


def test_header_found_on_other_sheet(tmp_path):
    wb = Workbook()
    wb.active.title = "안내"
    wb.active.append(["이 시트는 설명입니다"])
    ws = wb.create_sheet("데이터")
    ws.append(["전표조건검색"])
    ws.append([])
    ws.append(["조회조건: 2026-09"])
    ws.append(EXPECTED_COLUMNS)
    rec = build_rows(n_vouchers=1)[0]
    ws.append([rec[h] for h in EXPECTED_COLUMNS])
    path = tmp_path / "multi.xlsx"
    wb.save(path)
    cands = find_header_candidates(path)
    assert [(c.sheet, c.header_row) for c in cands] == [("데이터", 4)]
    res = load_vouchers(path, cands[0])
    assert not res.errors and len(res.df) == 1
    assert res.df["_엑셀행"].iloc[0] == 5


def test_unreadable_file_korean_message():
    with pytest.raises(VoucherFileError, match="엑셀 파일을 읽을 수 없습니다"):
        find_header_candidates(b"this is not an excel file")


def test_missing_required_column(tmp_path):
    headers = [h for h in EXPECTED_COLUMNS if h != "계정과목코드"]
    path = write_xlsx(tmp_path / "m.xlsx", headers=headers)
    res = load_vouchers(path, find_header_candidates(path)[0])
    assert any("필수 열이 없습니다" in e and "계정과목코드" in e for e in res.errors)


def test_duplicate_required_header_is_error(tmp_path):
    headers = EXPECTED_COLUMNS + ["차변금액"]
    path = write_xlsx(tmp_path / "d.xlsx", headers=headers)
    res = load_vouchers(path, find_header_candidates(path)[0])
    assert any("중복" in e and "차변금액" in e for e in res.errors)


def test_codes_kept_as_text_and_raw_preserved(tmp_path):
    rows = build_rows(n_vouchers=1)
    rows[0]["계정과목코드"] = "00123"
    rows[0]["거래처사업자번호"] = "0012345678"
    rows[0]["적요"] = "  원재료   매입  "
    rows[0]["기표번호"] = 20260901  # 숫자로 저장된 번호
    path = write_xlsx(tmp_path / "c.xlsx", rows=rows)
    res = load_vouchers(path, find_header_candidates(path)[0])
    r = res.df.iloc[0]
    assert r["계정과목코드"] == "00123"
    assert r["거래처사업자번호"] == "0012345678"
    assert r["기표번호"] == "20260901"
    assert r["적요"] == "원재료 매입"
    assert res.raw.iloc[0]["적요"] == "  원재료   매입  "
    assert isinstance(r["회계일"], dt.date)


def test_bad_amount_flagged_not_summed(tmp_path):
    rows = build_rows(n_vouchers=2)
    rows[0]["차변금액"] = "12,3a4"
    path = write_xlsx(tmp_path / "b.xlsx", rows=rows)
    res = load_vouchers(path, find_header_candidates(path)[0])
    first = res.df.iloc[0]
    assert first["차변금액"] is None and first["차변금액_상태"] == "오류"
    assert with_side(res.df).iloc[0]["차대구분"] == SIDE_ERROR
    assert any("해석할 수 없는" in w for w in res.warnings)


def test_voucher_key_uses_accounting_unit(tmp_path):
    rows = build_rows(n_vouchers=1)
    other = [dict(r, 회계단위="공장") for r in rows]
    path = write_xlsx(tmp_path / "k.xlsx", rows=rows + other)
    res = load_vouchers(path, find_header_candidates(path)[0])
    assert res.df["기표번호"].nunique() == 1
    assert res.df["전표키"].nunique() == 2


def test_source_file_not_modified(synthetic_path):
    before = hashlib.sha256(synthetic_path.read_bytes()).hexdigest()
    load_vouchers(synthetic_path, find_header_candidates(synthetic_path)[0])
    assert hashlib.sha256(synthetic_path.read_bytes()).hexdigest() == before


def test_flag_string_false_distinguished(loaded):
    assert set(loaded.df["파일첨부여부_구분"]) == {"문자열 FALSE"}
