"""제조업 전표 차변/대변 계정과목 질문 응답 엔진.

외부 AI/API 없이 로컬에서 동작한다.
1) 제조업 거래유형 가이드(JSON)에서 질문과 맞는 거래유형을 키워드로 찾고
2) 업로드한 회사 전표에서 해당 계정의 실제 사용 현황(코드·차대 빈도·상대계정)과
   비슷한 적요로 과거에 사용된 계정을 함께 보여준다.
답변은 참고용이며 회사 승인 계정 규칙이 아니다.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

from .accounts import SIDE_CREDIT, SIDE_DEBIT, counterpart_accounts, with_side

GUIDE_PATH = Path(__file__).parent / "data" / "manufacturing_guide.json"

# 질문에서 검색어로 쓰지 않는 말
STOPWORDS = {
    "차변", "대변", "계정", "계정과목", "과목", "전표", "분개", "처리", "입력", "작성", "제조업", "우리", "회사",
    "어떤", "어떻게", "무엇", "뭐", "뭘", "뭔가요", "써요", "쓰나요", "써야", "쓰면", "하나요", "해야", "하면",
    "알려줘", "알려주세요", "경우", "할때", "때", "관련", "질문", "건", "것", "좀", "무슨", "되나요", "돼요",
}
# 여러 거래에 두루 쓰이는 약한 키워드 (가중치 절반)
GENERIC_KEYWORDS = {
    "지급", "결제", "출금", "송금", "입금", "매출", "판매", "수입", "투입", "정산", "카드", "교육", "개발",
    "연구", "보수", "수리", "전화", "도서", "임시", "가수", "선수", "회비", "상각", "폐기", "불량", "선적",
}
_JOSA = ("에서는", "으로는", "에서", "으로", "에는", "하고", "이고", "하면", "할때", "을", "를", "이", "가", "은",
         "는", "에", "의", "로", "도", "와", "과", "만")


def norm(text: str) -> str:
    return re.sub(r"\s+", "", str(text or "")).lower()


def base_name(name: str) -> str:
    """'복리후생비(제조)' → '복리후생비', '운반비-판' → '운반비'."""
    n = norm(name)
    n = re.sub(r"[\(\[（].*?[\)\]）]$", "", n)
    n = re.split(r"[-_/]", n)[0]
    return n


def load_guide(path: str | Path | None = None) -> dict:
    data = json.loads(Path(path or GUIDE_PATH).read_text(encoding="utf-8"))
    validate_guide(data)
    return data


def validate_guide(data: dict) -> None:
    if not isinstance(data, dict) or not isinstance(data.get("scenarios"), list):
        raise ValueError("가이드 JSON 에 scenarios 목록이 없습니다.")
    for i, s in enumerate(data["scenarios"], start=1):
        for key in ("id", "거래유형", "키워드", "분개"):
            if key not in s:
                raise ValueError(f"{i}번째 거래유형에 '{key}' 항목이 없습니다.")
        for e in s["분개"]:
            if not e.get("차변") or not e.get("대변"):
                raise ValueError(f"거래유형 '{s['거래유형']}' 의 분개에 차변/대변이 모두 있어야 합니다.")


# ------------------------------------------------------------------ 결과 모델 ----
@dataclass
class ScenarioMatch:
    scenario: dict
    score: float
    matched_keywords: list[str]
    preferred_entries: list[int]  # 질문 맥락(제조/판관)에 맞는 분개 인덱스


@dataclass
class Answer:
    question: str
    context: str | None  # "제조" | "판관" | None
    side_focus: str | None  # "차변" | "대변" | None
    matches: list[ScenarioMatch] = field(default_factory=list)
    company_accounts: dict[str, pd.DataFrame] = field(default_factory=dict)  # 가이드 계정명 → 회사 계정 후보
    similar_memo: pd.DataFrame | None = None  # 비슷한 적요 전표에서 쓰인 계정
    similar_memo_terms: list[str] = field(default_factory=list)
    mentioned_accounts: list[dict] = field(default_factory=list)  # 질문에 직접 언급된 회사 계정 프로필
    notes: list[str] = field(default_factory=list)

    @property
    def found(self) -> bool:
        return bool(self.matches or self.mentioned_accounts or (self.similar_memo is not None and len(self.similar_memo)))


# ------------------------------------------------------------------ 엔진 ----
class AccountAdvisor:
    def __init__(self, guide: dict | None = None, vouchers: pd.DataFrame | None = None,
                 inventory: pd.DataFrame | None = None):
        self.guide = guide or load_guide()
        self.vouchers = with_side(vouchers) if vouchers is not None and len(vouchers) else None
        self.inventory = inventory
        self.aliases = {norm(k): {norm(a) for a in v} for k, v in self.guide.get("aliases", {}).items()}
        names = {norm(i["계정"]) for s in self.guide["scenarios"] for e in s["분개"] for i in e["차변"] + e["대변"]}
        names |= set(self.aliases) | {a for v in self.aliases.values() for a in v}
        if inventory is not None and len(inventory):
            names |= {norm(n) for n in inventory["계정과목"]} | {base_name(n) for n in inventory["계정과목"]}
        self.account_names = sorted((n for n in names if len(n) >= 2), key=len, reverse=True)

    # --- 질문 해석 ---
    def detect_context(self, q: str) -> str | None:
        ctx = self.guide.get("context_keywords", {})
        hits = {k: sum(1 for w in words if norm(w) in q) for k, words in ctx.items()}
        best = [k for k, v in hits.items() if v and v == max(hits.values())]
        return best[0] if len(best) == 1 else None

    @staticmethod
    def detect_side(q: str) -> str | None:
        d, c = "차변" in q, "대변" in q
        if d and not c:
            return SIDE_DEBIT
        if c and not d:
            return SIDE_CREDIT
        return None

    def _account_spans(self, q_norm: str) -> list[tuple[int, int]]:
        spans = []
        for n in self.account_names:
            for m in re.finditer(re.escape(n), q_norm):
                spans.append((m.start(), m.end()))
        return spans

    def _keyword_in(self, q_norm: str, kw: str, spans: list[tuple[int, int]] | None = None) -> bool:
        """키워드 일치. 질문 속 더 긴 계정명의 일부로만 나타나면 제외 (예: '외상매입금' 안의 '입금')."""
        k = norm(kw)
        if not k:
            return False
        if re.fullmatch(r"[a-z0-9/&]+", k):  # 영문 약어는 단어 경계로만
            return re.search(rf"(?<![a-z0-9]){re.escape(k)}(?![a-z0-9])", q_norm) is not None
        spans = self._account_spans(q_norm) if spans is None else spans
        for m in re.finditer(re.escape(k), q_norm):
            a, b = m.start(), m.end()
            inside = any(s <= a and b <= e and (e - s) > (b - a) for s, e in spans)
            if not inside:
                return True
        return False

    def match_scenarios(self, question: str, top_n: int = 3) -> list[ScenarioMatch]:
        q = norm(question)
        context = self.detect_context(q)
        spans = self._account_spans(q)
        results = []
        for s in self.guide["scenarios"]:
            hits = [k for k in s["키워드"] if self._keyword_in(q, k, spans)]
            # 더 긴 일치 키워드에 포함된 짧은 키워드는 중복 가산하지 않음
            hits = [k for k in hits if not any(norm(k) != norm(o) and norm(k) in norm(o) for o in hits)]
            if not hits:
                continue
            score = sum(len(norm(k)) * (0.5 if k in GENERIC_KEYWORDS else 1.0) for k in hits)
            if norm(s["거래유형"]) in q:
                score += 5
            preferred = []
            if context:
                preferred = [i for i, e in enumerate(s["분개"]) if context in e.get("조건", "")]
            results.append(ScenarioMatch(s, score, hits, preferred))
        results.sort(key=lambda m: (-m.score, m.scenario["id"]))
        if results:
            cutoff = results[0].score * 0.4
            results = [m for m in results if m.score >= cutoff]
        return results[:top_n]

    def query_terms(self, question: str, matches: list[ScenarioMatch]) -> list[str]:
        terms = []
        for m in matches:
            terms.extend(m.matched_keywords)
        for word in re.split(r"[\s,.?!·/()\"']+", question):
            w = word.strip()
            for j in _JOSA:
                if w.endswith(j) and len(w) - len(j) >= 2:
                    w = w[: -len(j)]
                    break
            if len(w) >= 2 and norm(w) not in STOPWORDS and w not in terms:
                terms.append(w)
        return list(dict.fromkeys(t for t in terms if norm(t) not in STOPWORDS))

    # --- 회사 데이터 조회 ---
    def company_candidates(self, guide_account: str) -> pd.DataFrame:
        """가이드 계정명과 일치/유사한 회사 계정 (파일에서 계산)."""
        cols = ["일치구분", "계정과목코드", "계정과목", "사용구분", "차변행수", "대변행수", "사용전표수", "비용구분"]
        if self.inventory is None or self.inventory.empty:
            return pd.DataFrame(columns=cols)
        target = norm(guide_account)
        alias = self.aliases.get(target, set())
        rows = []
        for r in self.inventory.to_dict("records"):
            cname, cbase = norm(r["계정과목"]), base_name(r["계정과목"])
            if cbase == target or cname == target:
                kind = "일치"
            elif cbase in alias or cname in alias:
                kind = "동의어"
            elif len(target) >= 3 and target in cname:
                kind = "유사"
            else:
                continue
            rows.append({"일치구분": kind, **{c: r.get(c) for c in cols[1:]}})
        order = {"일치": 0, "동의어": 1, "유사": 2}
        return (pd.DataFrame(rows, columns=cols)
                .sort_values(["일치구분", "사용전표수"], key=lambda s: s.map(order) if s.name == "일치구분" else -s)
                .reset_index(drop=True))

    def similar_memo_accounts(self, terms: list[str], limit: int = 15) -> pd.DataFrame:
        """적요에 검색어가 들어간 과거 전표에서 차변/대변에 쓰인 계정."""
        cols = ["차대구분", "계정과목코드", "계정과목", "전표수", "적요예시"]
        if self.vouchers is None or not terms:
            return pd.DataFrame(columns=cols)
        v = self.vouchers
        memo = v["적요"].astype(str).map(norm)
        hits = pd.Series(0, index=v.index)
        for t in terms:
            hits += memo.str.contains(re.escape(norm(t)), regex=True).astype(int)
        if hits.max() == 0:
            return pd.DataFrame(columns=cols)
        # 검색어가 가장 많이 들어간 적요만 사용 (흔한 단어 하나로 무관한 전표가 섞이지 않게)
        mask = hits == hits.max()
        keys = v.loc[mask, "전표키"].unique()
        sub = v[v["전표키"].isin(keys) & v["차대구분"].isin([SIDE_DEBIT, SIDE_CREDIT])]
        rows = []
        for (side, code, name), g in sub.groupby(["차대구분", "계정과목코드", "계정과목"]):
            examples = g.loc[mask.reindex(g.index, fill_value=False), "적요"]
            if examples.empty:
                examples = g["적요"]
            rows.append({"차대구분": side, "계정과목코드": code, "계정과목": name,
                         "전표수": int(g["전표키"].nunique()),
                         "적요예시": " / ".join(examples.drop_duplicates().head(2))})
        out = pd.DataFrame(rows, columns=cols)
        out["_o"] = out["차대구분"].map({SIDE_DEBIT: 0, SIDE_CREDIT: 1})
        return (out.sort_values(["_o", "전표수", "계정과목코드"], ascending=[True, False, True])
                .drop(columns="_o").groupby("차대구분", sort=False).head(limit).reset_index(drop=True))

    def mentioned_company_accounts(self, question: str) -> list[dict]:
        if self.inventory is None or self.inventory.empty:
            return []
        q = norm(question)
        found, seen = [], set()
        # 긴 이름부터 확인해 '부가세대급금' 이 '대급금' 보다 먼저 잡히게 함
        inv = self.inventory.assign(_len=self.inventory["계정과목"].map(lambda n: len(norm(n))))
        for r in inv.sort_values("_len", ascending=False).to_dict("records"):
            b = base_name(r["계정과목"])
            if len(b) < 2 or (norm(r["계정과목"]) not in q and b not in q):
                continue
            if any(b in s for s in seen):
                continue
            seen.add(b)
            prof = {k: v for k, v in r.items() if k != "_len"}
            if self.vouchers is not None:
                prof["차변일때_상대계정"] = counterpart_accounts(self.vouchers, r["계정과목코드"], SIDE_DEBIT).head(5)
                prof["대변일때_상대계정"] = counterpart_accounts(self.vouchers, r["계정과목코드"], SIDE_CREDIT).head(5)
            found.append(prof)
        return found[:5]

    # --- 답변 ---
    def ask(self, question: str) -> Answer:
        question = (question or "").strip()
        q = norm(question)
        ans = Answer(question, self.detect_context(q), self.detect_side(q))
        if not question:
            ans.notes.append("질문을 입력해 주세요. 예: '공장 전기요금 낼 때 차변 대변 뭐 써요?'")
            return ans
        ans.matches = self.match_scenarios(question)
        for m in ans.matches:
            for e in m.scenario["분개"]:
                for item in e["차변"] + e["대변"]:
                    name = item["계정"]
                    if name not in ans.company_accounts and not name.startswith("해당 "):
                        ans.company_accounts[name] = self.company_candidates(name)
        ans.similar_memo_terms = self.query_terms(question, ans.matches)
        ans.similar_memo = self.similar_memo_accounts(ans.similar_memo_terms)
        ans.mentioned_accounts = self.mentioned_company_accounts(question)

        if self.inventory is None:
            ans.notes.append("전표 파일을 불러오지 않아 회사 실제 사용 계정은 표시하지 않았습니다.")
        if ans.matches and any(len(m.scenario["분개"]) > 1 for m in ans.matches) and ans.context is None:
            ans.notes.append(
                "제조(공장·생산) 비용인지 판관(본사·영업·관리) 비용인지에 따라 계정이 달라질 수 있습니다. "
                "질문에 '공장' 또는 '본사' 등을 넣으면 해당 분개를 먼저 보여 드립니다."
            )
        if not ans.found:
            ans.notes.append(
                "일치하는 거래유형이나 과거 적요를 찾지 못했습니다. 거래 내용을 구체적으로 적어 주세요 "
                "(예: '생산라인 외주 도금 가공비', '거래처 명절 선물'). 회계팀 확인을 권장합니다."
            )
        return ans


def list_examples() -> list[str]:
    return [
        "공장 전기요금 낼 때 차변 대변 뭐 써요?",
        "원재료 외상으로 샀어요",
        "생산직 급여 계정 알려줘",
        "거래처 접대 식사 법인카드",
        "외주 가공비 세금계산서 받았어요",
        "제품 수출 매출 영세율",
        "기계장치 설비 구입",
        "제품 판매 택배 운반비",
    ]
