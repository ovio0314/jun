"""Streamlit 화면 기본 동작 (AppTest: 실제 app.py 를 헤드리스로 실행)."""

from pathlib import Path

import pytest
from streamlit.testing.v1 import AppTest

from make_synthetic import write_xlsx

APP = str(Path(__file__).resolve().parents[1] / "app.py")


@pytest.fixture
def app(tmp_path, monkeypatch):
    inp = tmp_path / "input"
    write_xlsx(inp / "합성_전표.xlsx")
    monkeypatch.setenv("VOUCHER_INPUT_DIRS", str(inp))
    monkeypatch.setenv("VOUCHER_LOCAL_DIR", str(tmp_path / "local"))
    at = AppTest.from_file(APP, default_timeout=60)
    at.run()
    assert not at.exception
    return at


def _button(at, label):
    return next(b for b in at.button if b.label == label)


def _load_and_check(at):
    sel = next(s for s in at.sidebar.selectbox if s.label == "또는 input 폴더 파일")
    sel.select("합성_전표.xlsx").run()
    _button(at, "불러오기").click().run()
    assert not at.exception
    _button(at, "검사 실행").click().run()
    assert not at.exception


def test_start_without_file(app):
    assert any("전표 검토 도우미" in t.value for t in app.title)
    assert any("검사 실행" in i.value for i in app.info)


def test_load_run_checks_and_summary(app):
    _load_and_check(app)
    metrics = {m.label: m.value for m in app.metric}
    total = int(metrics["전표 수"].replace(",", ""))
    parts = sum(int(metrics[k]) for k in ("🔴 오류", "🟠 확인 필요", "⚪ 미검사", "🟢 설정된 검사 통과"))
    assert parts == total == 60
    assert metrics["⚪ 미검사"] == "60"  # 기본 설정: 증빙·계정 규칙 없음 → 미검사, 통과로 바뀌지 않음
    assert any("다운로드 건수 대조" in c.value for c in app.caption)


def test_review_save_does_not_depend_on_checks(app):
    _load_and_check(app)
    state = app.session_state
    assert {r.검토상태 for r in state["reviews"].values()} == {"미검토"}


def test_reset_clears_session_only(app, tmp_path):
    _load_and_check(app)
    _button(app, "초기화 (현재 세션만 비우기)").click().run()
    assert not app.exception
    assert (tmp_path / "input" / "합성_전표.xlsx").exists()
    assert app.session_state["result"] is None
