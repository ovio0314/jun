import json

import pytest

from voucher.qa import GUIDE_PATH, AccountAdvisor, base_name, load_guide, validate_guide


@pytest.fixture
def plain():
    return AccountAdvisor()


@pytest.fixture
def advisor(loaded, inventory):
    return AccountAdvisor(vouchers=loaded.df, inventory=inventory)


def _entry_accounts(match, idx, side):
    return [i["계정"] for i in match.scenario["분개"][idx][side]]


def test_factory_electricity(plain):
    ans = plain.ask("공장 전기요금 낼 때 차변 대변 뭐 써요?")
    top = ans.matches[0]
    assert top.scenario["id"] == "ELECTRICITY"
    assert ans.context == "제조"
    assert "전력비" in _entry_accounts(top, top.preferred_entries[0], "차변")


def test_office_electricity_prefers_sga(plain):
    ans = plain.ask("본사 사무실 전기요금")
    top = ans.matches[0]
    assert ans.context == "판관"
    assert "수도광열비" in _entry_accounts(top, top.preferred_entries[0], "차변")


def test_account_name_substring_not_keyword(plain):
    # '외상매입금' 안의 '입금' 때문에 매출채권 회수로 오인하지 않아야 함
    ans = plain.ask("외상매입금 대변에 쓰나요")
    assert "AR_COLLECTION" not in [m.scenario["id"] for m in ans.matches]
    assert ans.side_focus == "대변"


def test_entertainment_and_alias(advisor):
    ans = advisor.ask("거래처 접대 식사 법인카드")
    ids = [m.scenario["id"] for m in ans.matches]
    assert "ENTERTAINMENT" in ids
    cand = ans.company_accounts["기업업무추진비"]
    assert list(cand["계정과목"]) == ["접대비"] and cand["일치구분"].iloc[0] == "동의어"


def test_similar_memo_is_narrow(advisor):
    ans = advisor.ask("거래처 접대 식사")
    debit = ans.similar_memo[ans.similar_memo["차대구분"] == "차변"]
    assert list(debit["계정과목"]) == ["접대비"]


def test_company_candidates_split_cost_center(advisor):
    c = advisor.company_candidates("복리후생비")
    assert set(c["계정과목"]) == {"복리후생비(제조)", "복리후생비(판)"}


def test_mentioned_account_profile(advisor):
    ans = advisor.ask("외상매입금은 언제 차변에 써요?")
    prof = ans.mentioned_accounts[0]
    assert prof["계정과목"] == "외상매입금"
    assert "보통예금" in set(prof["차변일때_상대계정"]["계정과목"])


def test_unknown_question(advisor):
    ans = advisor.ask("우주선 발사")
    assert not ans.found and any("찾지 못했습니다" in n for n in ans.notes)


def test_empty_question(plain):
    assert plain.ask("   ").notes


def test_no_file_note(plain):
    assert any("불러오지 않아" in n for n in plain.ask("원재료 매입").notes)


def test_guide_integrity():
    guide = load_guide()
    ids = [s["id"] for s in guide["scenarios"]]
    assert len(ids) == len(set(ids))
    for s in guide["scenarios"]:
        assert s["키워드"]
        for e in s["분개"]:
            assert e["차변"] and e["대변"]


def test_every_scenario_reachable_by_its_keywords(plain):
    for s in load_guide()["scenarios"]:
        ids = [m.scenario["id"] for m in plain.match_scenarios(" ".join(s["키워드"][:2]))]
        assert s["id"] in ids, s["id"]


def test_validate_guide_rejects_bad():
    with pytest.raises(ValueError):
        validate_guide({"scenarios": [{"id": "X", "거래유형": "x", "키워드": [], "분개": [{"차변": [], "대변": []}]}]})
    with pytest.raises(ValueError):
        validate_guide({})


def test_base_name():
    assert base_name("복리후생비(제조)") == "복리후생비"
    assert base_name("운반비-판") == "운반비"
    assert base_name(" 전 력 비 ") == "전력비"


def test_guide_file_is_valid_json():
    json.loads(GUIDE_PATH.read_text(encoding="utf-8"))
