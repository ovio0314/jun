"""세션 상태 묶음. 화면 코드가 st.session_state 키를 직접 흩뿌리지 않게 한다."""

from __future__ import annotations

import datetime as dt

import streamlit as st

from voucher.accounts import account_inventory
from voucher.checks import run_checks, voucher_summary
from voucher.evidence import ManualEvidenceStore
from voucher.review import content_hashes, init_reviews
from voucher.settings import settings_version

_DEFAULTS = {
    "result": None, "source_name": "", "source_sha256": "", "inventory": None,
    "results": None, "summary": None, "checked_version": "", "checked_at": "",
    "reviews": {}, "reviewer": "",
}


class AppState:
    def __init__(self, settings_loader):
        ss = st.session_state
        for k, v in _DEFAULTS.items():
            ss.setdefault(k, v if not isinstance(v, dict) else {})
        if "settings" not in ss:
            ss["settings"] = settings_loader()
        if "evidence" not in ss:
            ss["evidence"] = ManualEvidenceStore()

    def __getattr__(self, name):
        return st.session_state[name]

    def __setattr__(self, name, value):
        st.session_state[name] = value

    # --- 동작
    def set_loaded(self, result, source_name: str, sha: str, previous_reviews=None) -> None:
        self.result = result
        self.source_name = source_name
        self.source_sha256 = sha
        self.inventory = account_inventory(result.df) if not result.errors else None
        self.results = None
        self.summary = None
        self.reviews = init_reviews(content_hashes(result.df), previous_reviews) if not result.errors else {}

    def rerun_checks(self) -> None:
        if self.result is None or self.result.errors:
            return
        self.results = run_checks(self.result.df, self.settings, self.evidence)
        self.summary = voucher_summary(self.result.df, self.results)
        self.checked_version = settings_version(self.settings)
        self.checked_at = dt.datetime.now().isoformat(timespec="seconds")

    def reset(self) -> None:
        """현재 세션만 비운다 (원본 파일·저장 데이터·설정 파일은 삭제하지 않음)."""
        keep = {"settings", "guide"}
        for k in list(st.session_state.keys()):
            if k not in keep:
                del st.session_state[k]
