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


def build_result_sheets(summary: pd.DataFrame, results: pd.DataFrame, raw: pd.DataFrame, df: pd.DataFrame,
                        reviews: pd.DataFrame, criteria: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """검사 결과 엑셀: 전표요약 / 검사상세 / 원본분개 / 적용기준. 사용자 메모 포함."""
    rv = reviews[["전표키", "검토상태", "메모", "검토자", "검토일시", "변경감지", "이전검토상태"]] if len(reviews) else \
        pd.DataFrame(columns=["전표키", "검토상태", "메모", "검토자", "검토일시", "변경감지", "이전검토상태"])
    s = summary.merge(rv, on="전표키", how="left", validate="one_to_one")
    s["검토상태"] = s["검토상태"].fillna("미검토")
    original = raw.copy()
    original.insert(0, "전표키", list(df["전표키"]))
    return {"전표요약": s, "검사상세": results, "원본분개": original, "적용기준": criteria}


def check_export_counts(sheets: dict[str, pd.DataFrame], n_rows: int, n_vouchers: int) -> list[str]:
    """다운로드 결과의 원본/요약 건수 대조. 문제가 없으면 빈 목록."""
    problems = []
    if len(sheets["원본분개"]) != n_rows:
        problems.append(f"원본분개 {len(sheets['원본분개'])}행 ≠ 원본 {n_rows}행")
    if len(sheets["전표요약"]) != n_vouchers:
        problems.append(f"전표요약 {len(sheets['전표요약'])}건 ≠ 전표 {n_vouchers}건")
    if sheets["전표요약"]["전표키"].duplicated().any():
        problems.append("전표요약에 중복 전표키")
    if set(sheets["검사상세"]["전표키"]) - set(sheets["전표요약"]["전표키"]):
        problems.append("검사상세에 요약에 없는 전표키")
    return problems
