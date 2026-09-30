from decimal import Decimal

from voucher.evidence import (
    APPROVAL_DOC,
    BY_MANUAL,
    CONFIRM_OK,
    CONFIRM_PENDING,
    LINK_CANDIDATE,
    LINK_CONFIRMED,
    LINK_FAILED,
    SOURCE_OCR,
    TAX_INVOICE,
    FieldValue,
    ManualEvidenceStore,
    NotConnectedProvider,
    auto_link,
    suggest_by_partner_amount,
)
from voucher.checks import AREA_EVIDENCE, PASS, UNCHECKED, run_checks
from voucher.settings import default_settings

KEY = None


def _key(loaded):
    return loaded.df["전표키"].iloc[0]


def _ev(loaded, store):
    res = run_checks(loaded.df, default_settings(), store)
    return res[(res["전표키"] == _key(loaded)) & (res["검사영역"] == AREA_EVIDENCE)]


def test_partner_amount_is_candidate_only(loaded):
    store = ManualEvidenceStore()
    store.add_doc(TAX_INVOICE, {"승인번호": "A", "공급자사업자번호": "0001234567", "공급가액": Decimal(1),
                                "세액": Decimal(0), "합계": Decimal(1)})
    links = suggest_by_partner_amount(store, _key(loaded), "0001234567", Decimal(1))
    assert [l.status for l in links] == [LINK_CANDIDATE]
    ev = _ev(loaded, store)
    assert set(ev["상태"]) == {UNCHECKED}
    assert ev["근거"].str.contains("후보만 있음").any()


def test_ocr_value_requires_manual_confirmation(loaded):
    store = ManualEvidenceStore()
    ocr = lambda v: FieldValue(v, source=SOURCE_OCR, page=1, region=[10, 20, 110, 40], confidence=0.9,
                               confirm_status=CONFIRM_PENDING)
    doc = store.add_doc(APPROVAL_DOC, {"문서번호": ocr("품의-9"), "승인상태": ocr("승인")}, source=SOURCE_OCR)
    store.link(_key(loaded), doc.doc_id, BY_MANUAL)
    ev = _ev(loaded, store)
    assert (ev[ev["검사코드"] == "E02"]["상태"] == UNCHECKED).all()
    assert doc.fields["문서번호"].page == 1 and doc.fields["문서번호"].region == [10, 20, 110, 40]
    for f in doc.fields.values():
        f.confirm_status = CONFIRM_OK
    ev = _ev(loaded, store)
    assert (ev[ev["검사코드"] == "E02"]["상태"] == PASS).all()


def test_recognition_failure_and_link_failure_unchecked(loaded):
    store = ManualEvidenceStore()
    doc = store.add_doc(APPROVAL_DOC, {"문서번호": "X", "승인상태": "승인"})
    doc.recognition_failed = True
    store.link(_key(loaded), doc.doc_id, BY_MANUAL)
    assert _ev(loaded, store)["근거"].str.contains("인식 실패").any()
    store2 = ManualEvidenceStore()
    link = store2.link(_key(loaded), "없는문서", BY_MANUAL)
    assert link.status == LINK_FAILED
    assert _ev(loaded, store2)["근거"].str.contains("연결 실패").any()


def test_file_name_alone_is_not_evidence(loaded):
    store = ManualEvidenceStore()
    doc = store.add_doc(APPROVAL_DOC, {}, file_name="품의서_최종_승인.pdf")
    store.link(_key(loaded), doc.doc_id, BY_MANUAL)
    ev = _ev(loaded, store)
    assert set(ev["상태"]) == {UNCHECKED}


def test_auto_link_only_by_exact_approval_id(loaded):
    store = ManualEvidenceStore()
    doc = store.add_doc(APPROVAL_DOC, {"문서번호": "D1", "승인상태": "승인"}, approval_doc_id="DAOU-77")
    other = store.add_doc(APPROVAL_DOC, {"문서번호": "D2", "승인상태": "승인"}, approval_doc_id="DAOU-78")
    keys = set(loaded.df["전표키"])
    made = auto_link(store, keys, {_key(loaded): "DAOU-77"})
    assert [(l.doc_id, l.status) for l in made] == [(doc.doc_id, LINK_CONFIRMED)]
    assert not store.links_for("없는키") and other.doc_id not in {l.doc_id for l in store.links}


def test_store_roundtrip():
    store = ManualEvidenceStore()
    doc = store.add_doc(TAX_INVOICE, {"승인번호": "1", "공급가액": Decimal("1000"), "세액": Decimal("100")})
    store.link("k", doc.doc_id)
    back = ManualEvidenceStore.loads(store.dumps())
    assert back.get_doc(doc.doc_id).value("공급가액") == Decimal("1000")
    assert back.links_for("k")[0].status == LINK_CONFIRMED


def test_not_connected_provider_is_inert():
    p = NotConnectedProvider("다우오피스")
    assert p.list_docs() == [] and p.get_doc("x") is None
