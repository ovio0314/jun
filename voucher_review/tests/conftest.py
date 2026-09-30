import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "sample_data"))

from make_synthetic import build_rows, write_xlsx  # noqa: E402
from voucher.accounts import account_inventory  # noqa: E402
from voucher.parser import find_header_candidates, load_vouchers  # noqa: E402


@pytest.fixture
def synthetic_path(tmp_path):
    return write_xlsx(tmp_path / "합성.xlsx")


@pytest.fixture
def loaded(synthetic_path):
    cand = find_header_candidates(synthetic_path)[0]
    return load_vouchers(synthetic_path, cand)


@pytest.fixture
def inventory(loaded):
    return account_inventory(loaded.df)


@pytest.fixture
def rows():
    return build_rows()
