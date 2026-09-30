"""화면 공통 도우미."""

from __future__ import annotations

from decimal import Decimal

import pandas as pd
import streamlit as st

MONEY_COLS = ("차변합계", "대변합계", "차이", "차변금액", "대변금액", "외화차변금액", "외화대변금액")

STATUS_ICON = {"오류": "🔴 오류", "확인 필요": "🟠 확인 필요", "미검사": "⚪ 미검사", "통과": "🟢 통과",
               "해당 없음": "➖ 해당 없음", "설정된 검사 통과": "🟢 설정된 검사 통과"}


def won(v) -> str:
    if isinstance(v, Decimal):
        return f"{v:,.0f}" if v == v.to_integral_value() else f"{v:,}"
    return "" if v is None else str(v)


def show_money(df: pd.DataFrame, cols=MONEY_COLS) -> pd.DataFrame:
    out = df.copy()
    for c in cols:
        if c in out:
            out[c] = out[c].map(won)
    return out


def icons(df: pd.DataFrame, cols) -> pd.DataFrame:
    out = df.copy()
    for c in cols:
        if c in out:
            out[c] = out[c].map(lambda s: STATUS_ICON.get(s, s))
    return out


def split_list(text) -> list[str]:
    """'a, b,c' → ['a','b','c'] (이미 리스트면 그대로)."""
    if isinstance(text, list):
        return [str(t).strip() for t in text if str(t).strip()]
    if text is None or (isinstance(text, float) and text != text):
        return []
    return [t.strip() for t in str(text).split(",") if t.strip()]


def join_list(values) -> str:
    return ", ".join(values or [])


def flash_and_rerun(message: str, kind: str = "success") -> None:
    """상태를 바꾼 뒤 화면 전체를 새 상태로 다시 그린다 (앞서 그린 탭·사이드바가 옛 값을 보이지 않게)."""
    st.session_state["_flash"] = (kind, message)
    st.rerun()


def show_flash() -> None:
    item = st.session_state.pop("_flash", None)
    if item:
        getattr(st, item[0])(item[1])
