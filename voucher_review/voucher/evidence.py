"""증빙 연결 인터페이스 (2차 확장 준비).

1차에서는 OCR·외부 API(다우오피스, 국세청 등)를 구현하지 않는다. 대신
- 증빙 원천값 모델(세금계산서 / 품의서 / 계정 해설)
- OCR 등 추출값의 원문 위치(페이지·영역)와 수동 확인 상태
- 전표 키 / 결재문서 ID 기준 연결과 수동 연결
- 증빙 공급자(Provider) 프로토콜
을 정의하고, 수기 입력(ManualEvidenceStore)만 동작하게 한다.

원칙
- 연결은 전표 키 또는 결재문서 ID 가 정확히 일치할 때만 '자동 연결(확정)'.
  거래처+금액 일치는 '후보'로만 제안하고 확정하지 않는다(수동 연결 필요).
- 파일명만으로 첨부 적정성을 판단하지 않는다(파일명은 참고 정보로만 보관).
- 증빙 미제공 / 연결 실패 / 인식 실패 / 수동 확인 전 추출값은 '미검사'로 남긴다.
"""

from __future__ import annotations

import datetime as dt
import json
import uuid
from dataclasses import asdict, dataclass, field
from decimal import Decimal
from pathlib import Path
from typing import Any, Protocol

# 증빙 유형
TAX_INVOICE = "세금계산서"
APPROVAL_DOC = "품의서"
ACCOUNT_NOTE = "계정해설"

# 값 출처
SOURCE_MANUAL = "수기입력"
SOURCE_OCR = "OCR"
SOURCE_API = "외부연동"
SOURCE_FILE = "파일가져오기"

# 추출값 수동 확인 상태
CONFIRM_PENDING = "미확인"
CONFIRM_OK = "확인"
CONFIRM_FIXED = "수정확인"
CONFIRM_FAILED = "인식실패"

# 연결 상태
LINK_CONFIRMED = "확정"
LINK_CANDIDATE = "후보"  # 수동 확정 전에는 검사에 쓰지 않음
LINK_FAILED = "연결실패"

# 연결 방식
BY_VOUCHER_KEY = "전표키"
BY_APPROVAL_ID = "결재문서ID"
BY_MANUAL = "수동"
BY_PARTNER_AMOUNT = "거래처+금액(후보 전용)"


@dataclass
class FieldValue:
    """증빙에서 얻은 값 하나. OCR 값은 원문 위치와 수동 확인 상태를 가진다."""

    value: Any
    source: str = SOURCE_MANUAL
    page: int | None = None  # 원문 페이지 (1부터)
    region: list[float] | None = None  # [x0, y0, x1, y1] 원문 영역
    confidence: float | None = None
    confirm_status: str = CONFIRM_OK  # 수기입력은 입력자가 확인한 값
    confirmed_by: str = ""
    confirmed_at: str = ""

    @property
    def usable(self) -> bool:
        """검사에 쓸 수 있는 값인지. OCR/연동 값은 수동 확인 후에만 사용."""
        if self.value in (None, ""):
            return False
        if self.confirm_status in (CONFIRM_PENDING, CONFIRM_FAILED):
            return False
        return True


def fv(value, **kw) -> FieldValue:
    return FieldValue(value, **kw)


@dataclass
class EvidenceDoc:
    """증빙 문서. 유형별 필드는 fields 에 FieldValue 로 저장."""

    doc_id: str
    doc_type: str  # 세금계산서 / 품의서 / 계정해설
    fields: dict[str, FieldValue]
    source: str = SOURCE_MANUAL
    file_name: str = ""  # 참고용. 파일명으로 적정성을 판단하지 않는다.
    approval_doc_id: str = ""  # 다우오피스 결재문서 ID (있으면)
    registered_by: str = ""
    registered_at: str = ""
    recognition_failed: bool = False

    def value(self, name: str):
        f = self.fields.get(name)
        return f.value if f and f.usable else None

    def unusable_fields(self, names: list[str]) -> list[str]:
        return [n for n in names if not (self.fields.get(n) and self.fields[n].usable)]


# 유형별 원천값 필드 정의
DOC_FIELDS: dict[str, list[str]] = {
    TAX_INVOICE: ["승인번호", "공급자사업자번호", "작성일", "공급가액", "세액", "합계", "출처"],
    APPROVAL_DOC: ["문서번호", "승인상태", "목적", "금액", "기간시작", "기간종료"],
    ACCOUNT_NOTE: ["계정코드", "적용기준", "문서근거"],
}
# 검사에 반드시 필요한 필드
CHECK_FIELDS: dict[str, list[str]] = {
    TAX_INVOICE: ["승인번호", "공급가액", "세액"],
    APPROVAL_DOC: ["문서번호", "승인상태"],
    ACCOUNT_NOTE: ["계정코드", "적용기준"],
}


@dataclass
class EvidenceLink:
    voucher_key: str
    doc_id: str
    method: str  # 전표키 / 결재문서ID / 수동 / 거래처+금액(후보 전용)
    status: str  # 확정 / 후보 / 연결실패
    linked_by: str = ""
    linked_at: str = ""
    note: str = ""


class EvidenceProvider(Protocol):
    """증빙 공급자 인터페이스. 2차에서 다우오피스·OCR·국세청 연동을 이 형태로 구현한다."""

    name: str

    def list_docs(self) -> list[EvidenceDoc]: ...

    def get_doc(self, doc_id: str) -> EvidenceDoc | None: ...


class NotConnectedProvider:
    """미구현 연동 자리표시자 (다우오피스·OCR 등). 호출해도 외부로 나가지 않는다."""

    def __init__(self, name: str):
        self.name = name

    def list_docs(self) -> list[EvidenceDoc]:
        return []

    def get_doc(self, doc_id: str) -> EvidenceDoc | None:
        return None


def _now() -> str:
    return dt.datetime.now().isoformat(timespec="seconds")


def _to_json(obj):
    if isinstance(obj, Decimal):
        return {"__decimal__": str(obj)}
    if isinstance(obj, (dt.date, dt.datetime)):
        return obj.isoformat()
    raise TypeError(type(obj))


def _from_json(d):
    if isinstance(d, dict) and set(d) == {"__decimal__"}:
        return Decimal(d["__decimal__"])
    return d


class ManualEvidenceStore:
    """수기 입력 증빙과 연결 정보 (메모리 + JSON 직렬화). EvidenceProvider 구현."""

    name = "수기입력"

    def __init__(self, docs: dict[str, EvidenceDoc] | None = None, links: list[EvidenceLink] | None = None):
        self.docs: dict[str, EvidenceDoc] = docs or {}
        self.links: list[EvidenceLink] = links or []

    # Provider
    def list_docs(self) -> list[EvidenceDoc]:
        return list(self.docs.values())

    def get_doc(self, doc_id: str) -> EvidenceDoc | None:
        return self.docs.get(doc_id)

    # 등록/연결
    def add_doc(self, doc_type: str, values: dict, *, user: str = "", file_name: str = "",
                approval_doc_id: str = "", source: str = SOURCE_MANUAL) -> EvidenceDoc:
        if doc_type not in DOC_FIELDS:
            raise ValueError(f"알 수 없는 증빙 유형: {doc_type}")
        fields = {}
        for k, v in values.items():
            fields[k] = v if isinstance(v, FieldValue) else FieldValue(
                v, source=source, confirm_status=CONFIRM_OK if source == SOURCE_MANUAL else CONFIRM_PENDING,
                confirmed_by=user if source == SOURCE_MANUAL else "", confirmed_at=_now() if source == SOURCE_MANUAL else "")
        doc = EvidenceDoc(f"EV-{uuid.uuid4().hex[:10]}", doc_type, fields, source, file_name, approval_doc_id,
                          user, _now())
        self.docs[doc.doc_id] = doc
        return doc

    def link(self, voucher_key: str, doc_id: str, method: str = BY_MANUAL, *, user: str = "", note: str = "") -> EvidenceLink:
        if method == BY_PARTNER_AMOUNT:
            status = LINK_CANDIDATE  # 거래처+금액만으로는 확정하지 않음
        elif doc_id not in self.docs:
            status = LINK_FAILED
        else:
            status = LINK_CONFIRMED
        link = EvidenceLink(voucher_key, doc_id, method, status, user, _now(), note)
        self.links = [l for l in self.links if not (l.voucher_key == voucher_key and l.doc_id == doc_id)]
        self.links.append(link)
        return link

    def confirm_candidate(self, voucher_key: str, doc_id: str, user: str) -> None:
        for l in self.links:
            if l.voucher_key == voucher_key and l.doc_id == doc_id and l.status == LINK_CANDIDATE:
                l.status, l.method, l.linked_by, l.linked_at = LINK_CONFIRMED, BY_MANUAL, user, _now()

    def unlink(self, voucher_key: str, doc_id: str) -> None:
        self.links = [l for l in self.links if not (l.voucher_key == voucher_key and l.doc_id == doc_id)]

    def links_for(self, voucher_key: str) -> list[EvidenceLink]:
        return [l for l in self.links if l.voucher_key == voucher_key]

    # 직렬화
    def to_dict(self) -> dict:
        return {"docs": [asdict(d) for d in self.docs.values()], "links": [asdict(l) for l in self.links]}

    @classmethod
    def from_dict(cls, data: dict) -> "ManualEvidenceStore":
        docs = {}
        for d in data.get("docs", []):
            fields = {k: FieldValue(**{**f, "value": _from_json(f["value"])}) for k, f in d["fields"].items()}
            docs[d["doc_id"]] = EvidenceDoc(**{**d, "fields": fields})
        links = [EvidenceLink(**l) for l in data.get("links", [])]
        return cls(docs, links)

    def dumps(self) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False, default=_to_json)

    @classmethod
    def loads(cls, text: str) -> "ManualEvidenceStore":
        return cls.from_dict(json.loads(text, object_hook=_from_json))


def auto_link(store: ManualEvidenceStore, voucher_keys: set[str], approval_ids: dict[str, str] | None = None) -> list[EvidenceLink]:
    """정확한 키 기준 자동 연결.

    - 증빙에 결재문서 ID 가 있고 전표의 결재문서 ID 와 정확히 같을 때만 확정.
    - approval_ids: 전표키 → 결재문서 ID (ERP/다우오피스에서 확인된 경우에만 전달)
    거래처·금액 유사성으로는 연결하지 않는다.
    """
    made = []
    by_approval = {v: k for k, v in (approval_ids or {}).items() if v}
    for doc in store.list_docs():
        if doc.approval_doc_id and doc.approval_doc_id in by_approval:
            key = by_approval[doc.approval_doc_id]
            if key in voucher_keys and not any(l.doc_id == doc.doc_id and l.voucher_key == key for l in store.links):
                made.append(store.link(key, doc.doc_id, BY_APPROVAL_ID, user="자동(결재문서ID 일치)"))
    return made


def suggest_by_partner_amount(store: ManualEvidenceStore, voucher_key: str, partner_bizno: str,
                              total: Decimal) -> list[EvidenceLink]:
    """거래처 사업자번호+합계 일치 세금계산서를 '후보'로만 제안 (확정 아님)."""
    out = []
    for doc in store.list_docs():
        if doc.doc_type != TAX_INVOICE:
            continue
        if doc.value("공급자사업자번호") == partner_bizno and doc.value("합계") == total:
            out.append(store.link(voucher_key, doc.doc_id, BY_PARTNER_AMOUNT, note="거래처+금액 일치 후보 — 수동 확인 필요"))
    return out
