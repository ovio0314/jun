from decimal import Decimal

from make_synthetic import build_rows, write_xlsx
from voucher.accounts import (
    SIDE_CREDIT,
    SIDE_DEBIT,
    account_inventory,
    code_name_conflicts,
    counterpart_accounts,
    file_overview,
    management_item_profile,
    voucher_balance,
    with_side,
)
from voucher.parser import find_header_candidates, load_vouchers


def _load(path):
    return load_vouchers(path, find_header_candidates(path)[0]).df


def test_inventory_covers_every_row(loaded, inventory):
    df = loaded.df
    total = inventory[["차변행수", "대변행수", "차대동시행수", "금액없음행수", "금액오류행수"]].sum().sum()
    assert total == len(df)
    assert set(inventory["계정과목코드"]) == set(df["계정과목코드"])


def test_inventory_sums_match_decimal_totals(loaded, inventory):
    df = loaded.df
    assert sum(inventory["차변합계"], Decimal(0)) == sum((v for v in df["차변금액"] if v), Decimal(0))
    assert sum(inventory["대변합계"], Decimal(0)) == sum((v for v in df["대변금액"] if v), Decimal(0))


def test_usage_classification(inventory):
    by = inventory.set_index("계정과목")
    assert by.loc["원재료", "사용구분"] == "차변 전용"
    assert by.loc["미지급금", "사용구분"] == "대변 전용"
    assert by.loc["외상매입금", "차변행수"] > 0 and by.loc["외상매입금", "대변행수"] > 0


def test_overview_is_computed_from_file(loaded):
    ov = file_overview(loaded.df)
    rows = build_rows()
    assert ov["원본행수"] == len(rows)
    assert ov["전표수"] == len({r["기표번호"] for r in rows})
    assert ov["차대변불일치전표수"] == 0
    assert str(ov["회계일_시작"]) == "2026-09-01"


def test_balance_mismatch_detected(tmp_path):
    rows = build_rows(n_vouchers=3)
    rows[0]["차변금액"] = rows[0]["차변금액"] + 1
    bal = voucher_balance(_load(write_xlsx(tmp_path / "x.xlsx", rows=rows)))
    diff = bal[bal["차이"] != 0]
    assert len(diff) == 1 and diff["차이"].iloc[0] == Decimal(1)


def test_fx_not_added_to_krw(tmp_path):
    rows = build_rows(n_vouchers=1)
    rows[0]["외화차변금액"] = 999
    bal = voucher_balance(_load(write_xlsx(tmp_path / "fx.xlsx", rows=rows)))
    assert bal["차이"].iloc[0] == 0


def test_negative_amount_is_not_error(tmp_path):
    rows = build_rows(n_vouchers=1)
    for r in rows:
        r["차변금액"], r["대변금액"] = -r["차변금액"], -r["대변금액"]
    df = with_side(_load(write_xlsx(tmp_path / "n.xlsx", rows=rows)))
    assert set(df["차대구분"]) <= {SIDE_DEBIT, SIDE_CREDIT}


def test_both_sides_on_one_row(tmp_path):
    rows = build_rows(n_vouchers=1)
    rows[0]["대변금액"] = 10
    inv = account_inventory(_load(write_xlsx(tmp_path / "b.xlsx", rows=rows)))
    assert inv["차대동시행수"].sum() == 1


def test_counterparts(loaded):
    cp = counterpart_accounts(loaded.df, "51600", SIDE_DEBIT)  # 전력비
    assert cp.iloc[0]["계정과목"] == "미지급금"
    cp = counterpart_accounts(loaded.df, "25100", SIDE_CREDIT)  # 외상매입금 대변 → 차변 상대
    assert {"원재료", "부가세대급금", "외주가공비"} <= set(cp["계정과목"])


def test_management_items_are_per_account(loaded):
    raw = management_item_profile(loaded.df, "15300")
    assert list(raw["관리항목"]) == ["관리항목2"]
    assert management_item_profile(loaded.df, "51600").empty


def test_code_name_conflicts(tmp_path):
    rows = build_rows(n_vouchers=12)  # 원재료 매입 전표 2건
    rows[0]["계정과목"] = "원재료(구)"
    c = code_name_conflicts(_load(write_xlsx(tmp_path / "c.xlsx", rows=rows)))
    assert "코드 하나에 계정명 여러 개" in set(c["유형"])
