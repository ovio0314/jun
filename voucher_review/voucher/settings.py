"""검사 설정. 기본값은 보수적으로: 회사가 설정하지 않은 기준으로는 검사하지 않는다(미검사)."""

from __future__ import annotations

import copy
import hashlib
import json
from pathlib import Path

DATA_DIR = Path(__file__).parent / "data"
EXAMPLE_RULES_PATH = DATA_DIR / "example_rules.json"

ROUNDING_METHODS = {"버림": "ROUND_DOWN", "반올림": "ROUND_HALF_UP", "올림": "ROUND_UP"}

DEFAULT_SETTINGS: dict = {
    "vat": {
        # 부가세 행 식별: 설정된 계정코드 + 계정명 키워드
        "account_codes": [],
        "name_keywords": ["부가세", "부가가치세"],
        # 계정코드 → 공급가액이 들어 있는 관리항목 열. 사용자가 확인한 매핑만 사용 (기본 없음 → 미검사)
        "supply_columns": {},
        # '증빙' 열 값 중 일반과세 내부 계산 점검 대상 (기본 없음 → 미검사)
        "general_evidence_values": [],
        # 일반과세 규칙으로 판정하지 않는 특수 유형 (해당 없음 처리). 부분 일치.
        "excluded_evidence_keywords": ["불공제", "영세", "면세", "수정", "분할", "계산서(면세)", "수입"],
        "rate": "0.1",
        "rounding": "",  # 버림/반올림/올림 중 선택. 빈 값 → 미검사
        "tolerance": "0",  # 원 단위 허용 차이
    },
    "memo": {
        "blank_status": "오류",  # 적요 빈칸 처리 상태
        "rules": [],  # 거래 유형별 필수 적요 정보 (사용자 설정 시에만)
    },
    "account_rules": [],  # 회사 승인 계정 규칙 (없으면 계정 검사 미검사)
    "evidence": {
        "require_tax_invoice_for_vat": True,
        "require_approval_doc": True,
    },
    "required_fields": ["기표번호", "회계일", "계정과목코드", "계정과목"],
}


def default_settings() -> dict:
    return copy.deepcopy(DEFAULT_SETTINGS)


def merge_defaults(data: dict) -> dict:
    out = default_settings()
    for k, v in (data or {}).items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k].update(v)
        else:
            out[k] = v
    return out


def settings_version(settings: dict) -> str:
    raw = json.dumps(settings, ensure_ascii=False, sort_keys=True, default=str)
    return "S-" + hashlib.sha256(raw.encode("utf-8")).hexdigest()[:10]


def load_settings(path: str | Path | None) -> dict:
    if path and Path(path).exists():
        return merge_defaults(json.loads(Path(path).read_text(encoding="utf-8")))
    return default_settings()


def save_settings(settings: dict, path: str | Path) -> None:
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    Path(path).write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")


def example_rules() -> dict:
    """합성 데이터용 예시 규칙. 모든 규칙이 비활성(활성=False)으로 제공된다."""
    return json.loads(EXAMPLE_RULES_PATH.read_text(encoding="utf-8"))


ACCOUNT_RULE_FIELDS = [
    "규칙ID", "이름", "적요조건", "적용전표유형", "차대구분", "허용계정코드", "필수상대계정코드",
    "근거", "적용시작일", "적용종료일", "활성", "버전", "예시",
]
MEMO_RULE_FIELDS = ["규칙ID", "이름", "대상계정코드", "적요조건", "필수포함", "근거", "활성", "버전", "예시"]
