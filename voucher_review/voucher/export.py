"""계정과목 현황 엑셀 내보내기. 텍스트 셀의 수식 주입을 방지한다."""

from __future__ import annotations

import io
from decimal import Decimal

import pandas as pd

_FORMULA_PREFIX = ("=", "+", "-", "@", "\t", "\r", "\n")


def safe_cell(value):
    """문자열이 수식 시작 문자로 시작하면 작은따옴표를 붙여 텍스트로 고정."""
    if isinstance(value, str) and value.startswith(_FORMULA_PREFIX):
        return "'" + value
    if isinstance(value, Decimal):
        # 원 단위 정수는 정수로, 그 외는 문자열로 보존해 이진 부동소수 오차를 피함
        return int(value) if value == value.to_integral_value() else str(value)
    return value


def safe_frame(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for c in out.columns:
        out[c] = out[c].map(safe_cell)
    out.columns = [safe_cell(str(c)) for c in out.columns]
    return out


def to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, df in sheets.items():
            safe_frame(df).to_excel(writer, sheet_name=name[:31], index=False)
    return buf.getvalue()
