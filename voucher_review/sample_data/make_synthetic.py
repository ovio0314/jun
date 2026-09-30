"""익명 합성 전표 엑셀 생성기 (테스트·시연용).

실제 회사 데이터가 아니며 계정코드·거래처·사람 이름은 모두 가상입니다.
실행: python sample_data/make_synthetic.py [출력경로]
"""

from __future__ import annotations

import datetime as dt
import random
import sys
from pathlib import Path

from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from voucher.columns import EXPECTED_COLUMNS  # noqa: E402

# 가상 계정표: 코드 → (계정명, 비용구분)
ACCOUNTS = {
    "10310": ("보통예금", ""),
    "10800": ("외상매출금", ""),
    "13500": ("부가세대급금", ""),
    "15300": ("원재료", ""),
    "20600": ("기계장치", ""),
    "25100": ("외상매입금", ""),
    "25300": ("미지급금", ""),
    "25400": ("예수금", ""),
    "25500": ("부가세예수금", ""),
    "40400": ("제품매출", ""),
    "50400": ("임금", "제조"),
    "51100": ("복리후생비(제조)", "제조"),
    "51600": ("전력비", "제조"),
    "53300": ("외주가공비", "제조"),
    "52000": ("수선비(제조)", "제조"),
    "80200": ("급여", "판관"),
    "81100": ("복리후생비(판)", "판관"),
    "81300": ("접대비", "판관"),
    "82400": ("운반비", "판관"),
    "81500": ("수도광열비", "판관"),
}

DEPTS = {"제조": ["생산1팀", "생산2팀", "품질팀"], "판관": ["영업팀", "경영지원팀"]}
WRITERS = ["직원A", "직원B", "직원C", "직원D", "직원E"]


def _line(code, debit=None, credit=None, memo="", **extra):
    return {"code": code, "debit": debit, "credit": credit, "memo": memo, **extra}


def _templates(rng: random.Random):
    """(적요, 부서구분, 분개 행) 생성 함수 목록."""

    def raw():
        s = rng.randrange(100, 3000) * 1000
        v = s // 10
        memo = f"원재료 매입 거래처{rng.randint(1, 9)}"
        return "제조", [_line("15300", s, memo=memo, m2=str(s)), _line("13500", v, memo=memo, m2=str(s)),
                        _line("25100", credit=s + v, memo=memo)]

    def power():
        s = rng.randrange(500, 5000) * 1000
        v = s // 10
        memo = "공장 전기요금"
        return "제조", [_line("51600", s, memo=memo), _line("13500", v, memo=memo), _line("25300", credit=s + v, memo=memo)]

    def office_power():
        s = rng.randrange(50, 300) * 1000
        v = s // 10
        memo = "본사 사무실 전기요금"
        return "판관", [_line("81500", s, memo=memo), _line("13500", v, memo=memo), _line("25300", credit=s + v, memo=memo)]

    def outsourcing():
        s = rng.randrange(200, 2000) * 1000
        v = s // 10
        memo = "외주 도금 가공비"
        return "제조", [_line("53300", s, memo=memo), _line("13500", v, memo=memo), _line("25100", credit=s + v, memo=memo)]

    def sales():
        s = rng.randrange(1000, 9000) * 1000
        v = s // 10
        memo = f"제품 매출 고객사{rng.randint(1, 9)}"
        return "판관", [_line("10800", s + v, memo=memo), _line("40400", credit=s, memo=memo),
                        _line("25500", credit=v, memo=memo)]

    def meal_factory():
        s = rng.randrange(100, 800) * 1000
        memo = "생산팀 야근 식대"
        return "제조", [_line("51100", s, memo=memo), _line("25300", credit=s, memo=memo)]

    def meal_office():
        s = rng.randrange(100, 500) * 1000
        memo = "영업팀 회식"
        return "판관", [_line("81100", s, memo=memo), _line("25300", credit=s, memo=memo)]

    def entertain():
        s = rng.randrange(100, 900) * 1000
        memo = "거래처 접대 식사"
        return "판관", [_line("81300", s, memo=memo), _line("25300", credit=s, memo=memo)]

    def payroll():
        g = rng.randrange(3000, 9000) * 10000
        w = g // 10
        return "제조", [_line("50400", g, memo="생산직 급여"), _line("25400", credit=w, memo="생산직 급여 원천세"),
                        _line("10310", credit=g - w, memo="생산직 급여")]

    def ap_pay():
        s = rng.randrange(100, 3000) * 1000
        memo = "외상매입금 지급"
        return "판관", [_line("25100", s, memo=memo), _line("10310", credit=s, memo=memo)]

    def freight():
        s = rng.randrange(30, 300) * 1000
        v = s // 10
        memo = "제품 택배 운반비"
        return "판관", [_line("82400", s, memo=memo), _line("13500", v, memo=memo), _line("25300", credit=s + v, memo=memo)]

    return [raw, power, office_power, outsourcing, sales, meal_factory, meal_office, entertain, payroll, ap_pay, freight]


def build_rows(n_vouchers: int = 60, seed: int = 7) -> list[dict]:
    rng = random.Random(seed)
    templates = _templates(rng)
    rows = []
    serial = 1
    for i in range(n_vouchers):
        kind, lines = templates[i % len(templates)]()
        date = dt.date(2026, 9, 1) + dt.timedelta(days=i % 30)
        no = f"A{date:%Y%m%d}{i + 1:04d}"
        dept = rng.choice(DEPTS[kind])
        writer = rng.choice(WRITERS)
        for ln, line in enumerate(lines, start=1):
            name, cost = ACCOUNTS[line["code"]]
            rec = {c: None for c in EXPECTED_COLUMNS}
            rec.update(
                회계단위="본사", 회계일=date, 전표분개유형="일반", 전표관리단위="회계", 기표번호=no,
                전표기표번호=f"{no}-{ln:03d}", 행번호=ln, 기표자=writer, 기표부서=dept, 기표일=date,
                승인자="승인자X", 승인부서="재무팀", 승인번호=f"AP{serial:05d}", 계정과목코드=line["code"],
                계정과목=name, 비용구분=cost, 차변금액=line["debit"] or 0, 대변금액=line["credit"] or 0,
                외화차변금액=0, 외화대변금액=0, 적요=line["memo"], 통화="KRW", 환율=1, 귀속부서=dept,
                증빙="세금계산서" if line["code"] == "13500" else "", 전표종류그룹="일반전표",
                거래처사업자번호="0001234567" if line["code"] in ("25100", "25300", "10800") else "",
                전자결재진행상태="결재완료", 관리항목2=line.get("m2"), 파일첨부여부="FALSE",
            )
            rows.append(rec)
            serial += 1
    return rows


def write_xlsx(path: str | Path, rows: list[dict] | None = None, title: str = "전표조건검색",
               headers: list[str] | None = None, sheet: str = "Sheet1") -> Path:
    rows = build_rows() if rows is None else rows
    headers = headers or EXPECTED_COLUMNS
    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    ws.append([title])
    ws.append(headers)
    for r in rows:
        ws.append([r.get(h) for h in headers])
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else str(Path(__file__).parent / "합성_전표_샘플.xlsx")
    print(write_xlsx(out))
