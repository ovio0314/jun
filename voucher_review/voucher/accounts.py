"""전표 데이터에서 차변/대변에 쓰인 계정과목을 모두 파악한다."""

from __future__ import annotations

from collections import Counter
from decimal import Decimal

import pandas as pd

from .columns import MANAGEMENT_ITEM_COLUMNS

ZERO = Decimal("0")

SIDE_DEBIT = "차변"
SIDE_CREDIT = "대변"
SIDE_BOTH = "차대동시"
SIDE_NONE = "금액없음"
SIDE_ERROR = "금액오류"


def _amount(value) -> Decimal:
    return value if isinstance(value, Decimal) else ZERO


def classify_side(row) -> str:
    """행 단위 차변/대변 판정.

    - 금액 해석 오류가 있으면 '금액오류' (차/대 어느 쪽으로도 집계하지 않음)
    - 빈 금액은 0 으로 본다(한쪽만 입력하는 ERP 관행). 상태는 *_상태 열에 별도 보존.
    - 음수도 입력된 쪽의 금액으로 본다(역분개/수정 가능성). 오류로 판정하지 않는다.
    """
    if row.get("차변금액_상태") == "오류" or row.get("대변금액_상태") == "오류":
        return SIDE_ERROR
    d, c = _amount(row.get("차변금액")), _amount(row.get("대변금액"))
    if d != ZERO and c != ZERO:
        return SIDE_BOTH
    if d != ZERO:
        return SIDE_DEBIT
    if c != ZERO:
        return SIDE_CREDIT
    return SIDE_NONE


def with_side(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["차대구분"] = [classify_side(r) for r in out.to_dict("records")] if len(out) else []
    return out


def _top(values, n=3) -> str:
    counter = Counter(v for v in values if v)
    return " / ".join(f"{k} ({c})" for k, c in counter.most_common(n))


def account_inventory(df: pd.DataFrame) -> pd.DataFrame:
    """계정과목코드+계정과목 별 차변/대변 사용 현황."""
    data = with_side(df)
    rows = []
    for (code, name), g in data.groupby(["계정과목코드", "계정과목"], sort=True, dropna=False):
        sides = g["차대구분"].value_counts()
        n_debit, n_credit = int(sides.get(SIDE_DEBIT, 0)), int(sides.get(SIDE_CREDIT, 0))
        debit_sum = sum((_amount(v) for v in g.loc[g["차대구분"].isin([SIDE_DEBIT, SIDE_BOTH]), "차변금액"]), ZERO)
        credit_sum = sum((_amount(v) for v in g.loc[g["차대구분"].isin([SIDE_CREDIT, SIDE_BOTH]), "대변금액"]), ZERO)
        if n_debit and not n_credit:
            usage = "차변 전용"
        elif n_credit and not n_debit:
            usage = "대변 전용"
        elif n_debit and n_credit:
            usage = "차변 위주" if n_debit >= n_credit * 3 else ("대변 위주" if n_credit >= n_debit * 3 else "양쪽 사용")
        else:
            usage = "판정 불가"
        rows.append(
            {
                "계정과목코드": code,
                "계정과목": name,
                "사용구분": usage,
                "차변행수": n_debit,
                "대변행수": n_credit,
                "차대동시행수": int(sides.get(SIDE_BOTH, 0)),
                "금액없음행수": int(sides.get(SIDE_NONE, 0)),
                "금액오류행수": int(sides.get(SIDE_ERROR, 0)),
                "차변합계": debit_sum,
                "대변합계": credit_sum,
                "사용전표수": int(g["전표키"].nunique()),
                "비용구분": _top(g.get("비용구분", []), 3),
                "기표부서(상위)": _top(g.get("기표부서", []), 3),
                "적요예시(상위)": _top(g.get("적요", []), 3),
            }
        )
    return pd.DataFrame(rows)


def code_name_conflicts(df: pd.DataFrame) -> pd.DataFrame:
    """하나의 코드에 여러 계정명, 또는 하나의 계정명에 여러 코드가 쓰인 경우."""
    out = []
    for code, names in df.groupby("계정과목코드")["계정과목"].unique().items():
        if len(names) > 1:
            out.append({"유형": "코드 하나에 계정명 여러 개", "기준": code, "값": ", ".join(sorted(names))})
    for name, codes in df.groupby("계정과목")["계정과목코드"].unique().items():
        if len(codes) > 1:
            out.append({"유형": "계정명 하나에 코드 여러 개", "기준": name, "값": ", ".join(sorted(codes))})
    return pd.DataFrame(out, columns=["유형", "기준", "값"])


def counterpart_accounts(df: pd.DataFrame, code: str, side: str) -> pd.DataFrame:
    """해당 계정이 `side` 쪽에 쓰인 전표에서 반대편에 함께 쓰인 계정 (전표 수 기준)."""
    data = with_side(df)
    opposite = SIDE_CREDIT if side == SIDE_DEBIT else SIDE_DEBIT
    keys = data.loc[(data["계정과목코드"] == code) & (data["차대구분"] == side), "전표키"].unique()
    sub = data[data["전표키"].isin(keys) & (data["차대구분"] == opposite) & (data["계정과목코드"] != code)]
    if sub.empty:
        return pd.DataFrame(columns=["계정과목코드", "계정과목", "함께쓴전표수"])
    res = (
        sub.groupby(["계정과목코드", "계정과목"])["전표키"].nunique().reset_index(name="함께쓴전표수")
        .sort_values(["함께쓴전표수", "계정과목코드"], ascending=[False, True])
    )
    return res.reset_index(drop=True)


def management_item_profile(df: pd.DataFrame, code: str) -> pd.DataFrame:
    """계정별 관리항목 열 사용 현황. 같은 열이라도 계정마다 의미가 다르므로 계정 단위로만 본다."""
    g = df[df["계정과목코드"] == code]
    rows = []
    for col in MANAGEMENT_ITEM_COLUMNS:
        if col not in g:
            continue
        filled = g[col].astype(str).str.len() > 0
        if not filled.any():
            continue
        rows.append(
            {
                "관리항목": col,
                "입력행수": int(filled.sum()),
                "전체행수": len(g),
                "예시값(상위)": _top(g.loc[filled, col], 3),
            }
        )
    return pd.DataFrame(rows, columns=["관리항목", "입력행수", "전체행수", "예시값(상위)"])


def voucher_balance(df: pd.DataFrame) -> pd.DataFrame:
    """전표키별 원화 차대변 합계. 외화 금액은 더하지 않는다."""
    data = with_side(df)
    rows = []
    for key, g in data.groupby("전표키", sort=True):
        ok = g["차대구분"] != SIDE_ERROR
        d = sum((_amount(v) for v in g.loc[ok, "차변금액"]), ZERO)
        c = sum((_amount(v) for v in g.loc[ok, "대변금액"]), ZERO)
        rows.append(
            {
                "전표키": key,
                "기표번호": g["기표번호"].iloc[0],
                "행수": len(g),
                "차변합계": d,
                "대변합계": c,
                "차이": d - c,
                "금액오류행": int((~ok).sum()),
            }
        )
    return pd.DataFrame(rows, columns=["전표키", "기표번호", "행수", "차변합계", "대변합계", "차이", "금액오류행"])


def file_overview(df: pd.DataFrame) -> dict:
    """화면 상단 요약. 값은 모두 파일에서 계산한다(하드코딩 금지)."""
    dates = [d for d in df.get("회계일", []) if d is not None]
    bal = voucher_balance(df) if len(df) else pd.DataFrame(columns=["차이"])
    inv = account_inventory(df) if len(df) else pd.DataFrame()
    flags = df["파일첨부여부_구분"].value_counts().to_dict() if "파일첨부여부_구분" in df else {}
    return {
        "원본행수": len(df),
        "전표수": int(df["전표키"].nunique()) if len(df) else 0,
        "기표번호수": int(df["기표번호"].nunique()) if len(df) else 0,
        "회계일_시작": min(dates) if dates else None,
        "회계일_종료": max(dates) if dates else None,
        "계정과목수": len(inv),
        "차변사용계정수": int((inv["차변행수"] > 0).sum()) if len(inv) else 0,
        "대변사용계정수": int((inv["대변행수"] > 0).sum()) if len(inv) else 0,
        "차대변불일치전표수": int((bal["차이"] != ZERO).sum()) if len(bal) else 0,
        "적요빈칸행수": int((df["적요"] == "").sum()) if "적요" in df and len(df) else 0,
        "파일첨부여부구분": flags,
    }
