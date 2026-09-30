"""브라우저판(JavaScript) 엔진이 파이썬 엔진과 같은 결과를 내는지 대조한다.

node 와 browser/node_modules 가 없으면 건너뛴다 (browser 폴더에서 npm install).
"""

import copy
import datetime as dt
import json
import shutil
import subprocess
from decimal import Decimal
from pathlib import Path

import pytest

from make_synthetic import build_rows, write_xlsx
from voucher.accounts import account_inventory, file_overview
from voucher.checks import run_checks, voucher_summary
from voucher.evidence import BY_MANUAL, ManualEvidenceStore
from voucher.parser import find_header_candidates, load_vouchers
from voucher.qa import AccountAdvisor
from voucher.review import content_hashes
from voucher.settings import default_settings, example_rules

BROWSER = Path(__file__).resolve().parents[1] / "browser"
pytestmark = pytest.mark.skipif(
    shutil.which("node") is None or not (BROWSER / "node_modules" / "exceljs").exists(),
    reason="node 또는 browser/node_modules 없음",
)
QUESTIONS = ["공장 전기요금 낼 때 차변 대변 뭐 써요?", "원재료 외상으로 샀어요", "외상매입금 대변에 쓰나요",
             "거래처 접대 식사 법인카드", "본사 사무실 전기요금", "우주선 발사"]


def plain(v):
    if isinstance(v, Decimal):
        return str(v)
    if isinstance(v, (dt.date, dt.datetime)):
        return v.isoformat()
    if hasattr(v, "item"):  # numpy 정수
        return v.item()
    return v


def records(df):
    return [{k: plain(v) for k, v in r.items()} for r in df.to_dict("records")]


def python_side(path, settings, evidence):
    res = load_vouchers(path, find_header_candidates(path)[0])
    store = None
    if evidence:
        store = ManualEvidenceStore()
        for e in evidence:
            vals = {k: Decimal(v) if k in ("공급가액", "세액", "합계", "금액") else v for k, v in e["values"].items()}
            doc = store.add_doc(e["type"], vals)
            store.link(e["key"], doc.doc_id, BY_MANUAL)
    results = run_checks(res.df, settings, store).drop(columns="적용규칙버전")
    inv = account_inventory(res.df)
    adv = AccountAdvisor(vouchers=res.df, inventory=inv)
    qa = []
    for q in QUESTIONS:
        a = adv.ask(q)
        qa.append({"context": a.context, "side": a.side_focus, "ids": [m.scenario["id"] for m in a.matches],
                   "preferred": [m.preferred_entries for m in a.matches], "terms": a.similar_memo_terms,
                   "similar": records(a.similar_memo) if a.similar_memo is not None else [],
                   "mentioned": [m["계정과목코드"] for m in a.mentioned_accounts]})
    ov = file_overview(res.df)
    return {
        "errors": res.errors, "warnings": res.warnings, "preamble": res.preamble, "header_row": res.header_row,
        "keys": list(res.df["전표키"]), "results": records(results),
        "summary": records(voucher_summary(res.df, run_checks(res.df, settings, store))),
        "overview": {k: plain(v) for k, v in ov.items()}, "inventory": records(inv), "qa": qa,
        "hashes": content_hashes(res.df),
    }


def js_side(tmp_path, path, settings, evidence):
    sp = tmp_path / "settings.json"
    sp.write_text(json.dumps(settings, ensure_ascii=False), encoding="utf-8")
    args = ["node", str(BROWSER / "test" / "run_engine.cjs"), str(path), str(sp)]
    if evidence:
        ep = tmp_path / "ev.json"
        ep.write_text(json.dumps(evidence, ensure_ascii=False), encoding="utf-8")
        args.append(str(ep))
    out = subprocess.run(args, capture_output=True, text=True, check=True, cwd=BROWSER, timeout=120)
    return json.loads(out.stdout)


def edge_rows():
    """경계 사례를 모은 합성 전표."""
    rows = build_rows(n_vouchers=22)
    rows[0]["차변금액"] = "1,000,O00"                      # 잘못된 숫자
    rows[3]["대변금액"] = f"{rows[3]['대변금액']:,}"        # 쉼표 숫자
    rows[6]["대변금액"] = 10                                # 차대 동시 입력
    rows[9]["차변금액"] += 7                                # 불일치
    for r in rows[12:15]:
        r["기표번호"] = ""                                  # 공백 번호
    rows.append(copy.deepcopy(rows[20]))                    # 완전 중복 행
    for r in rows[16:19]:
        r["차변금액"], r["대변금액"] = -r["차변금액"], -r["대변금액"]  # 음수
    vat_rows = [r for r in rows[20:] if r["계정과목코드"] == "13500"]
    vat_rows[0]["증빙"] = "영세율세금계산서"                  # 특수 증빙 (부가세 행)
    rows[25]["적요"] = ""                                   # 적요 빈칸
    other = [dict(r, 회계단위="공장") for r in build_rows(n_vouchers=2)]  # 복수 회계단위 동일 번호
    other[0]["차변금액"] += 1
    return rows + other


def configured():
    s = default_settings()
    ex = example_rules()
    s["account_rules"] = [dict(r, 활성=True) for r in ex["account_rules"]]
    s["memo"]["rules"] = [dict(r, 활성=True) for r in ex["memo_rules"]]
    s["vat"].update(ex["vat"])
    return s


CASES = {
    "synthetic_default": (lambda: build_rows(), default_settings, None),
    "synthetic_configured": (lambda: build_rows(), configured, None),
    "edge_default": (edge_rows, default_settings, None),
    "edge_configured": (edge_rows, configured, None),
    "evidence": (lambda: build_rows(n_vouchers=3), configured, [
        {"key": "본사|회계|A202609010001", "type": "세금계산서", "values": {"승인번호": "T1", "공급가액": "1426000", "세액": "142600", "합계": "1568600"}},
        {"key": "본사|회계|A202609010001", "type": "품의서", "values": {"문서번호": "D1", "승인상태": "승인"}},
        {"key": "본사|회계|A202609020002", "type": "품의서", "values": {"문서번호": "D2", "승인상태": "반려"}},
    ]),
}


@pytest.mark.parametrize("case", list(CASES))
def test_js_engine_matches_python(tmp_path, case):
    make_rows, make_settings, evidence = CASES[case]
    path = write_xlsx(tmp_path / f"{case}.xlsx", rows=make_rows())
    settings = make_settings()
    py = python_side(path, settings, evidence)
    js = js_side(tmp_path, path, settings, evidence)
    for part in ("errors", "warnings", "preamble", "header_row", "keys", "overview", "inventory", "hashes", "qa"):
        assert js[part] == py[part], part
    assert len(js["results"]) == len(py["results"])
    for a, b in zip(js["results"], py["results"]):
        assert a == b
    assert js["summary"] == py["summary"]
    if case == "edge_configured":  # 경계 사례가 실제로 비교되었는지 확인
        seen = {(r["검사코드"], r["상태"]) for r in js["results"]}
        for need in [("D01", "오류"), ("D02", "확인 필요"), ("D03", "확인 필요"), ("D04", "오류"), ("D05", "오류"),
                     ("A01", "오류"), ("A01", "미검사"), ("V01", "통과"), ("V01", "해당 없음"), ("M01", "오류")]:
            assert need in seen, need


def test_built_html_is_up_to_date():
    """커밋된 dist/전표검토.html 이 현재 src 로 빌드된 것인지 확인 (npm run build 누락 방지)."""
    html = (BROWSER / "dist" / "전표검토.html").read_text(encoding="utf-8")
    for src in ("src/engine.js", "src/app.js", "src/style.css"):
        text = (BROWSER / src).read_text(encoding="utf-8").replace("</script", "<\\/script").replace("<!--", "<\\!--")
        assert text in html, f"{src} 가 빌드에 반영되지 않았습니다. browser 폴더에서 npm run build 를 실행하세요."
    assert "connect-src 'none'" in html and "unsafe-eval" not in html
