"""사용자 검토 상태 · 로컬 저장/재열기.

- 검토상태(미검토/검토중/검토완료)·메모·검토자·검토일시는 자동 검사 결과와 분리해 저장한다.
- 검사 실행으로 검토상태가 바뀌지 않는다.
- 전표 내용 해시로 분개 변경을 감지한다. 같은 전표 키라도 내용이 바뀌면
  기존 '검토완료'를 승계하지 않고 '미검토' + 변경 표시로 되돌린다.
"""

from __future__ import annotations

import datetime as dt
import hashlib
import json
import re
from dataclasses import asdict, dataclass
from decimal import Decimal
from pathlib import Path

import pandas as pd

from .columns import EXPECTED_COLUMNS

REVIEW_STATES = ["미검토", "검토중", "검토완료"]
SAVE_FORMAT = "voucher-review-save/1"

# 내용 해시에 쓰는 열 (전표 키와 행 순서를 제외한 분개 내용)
HASH_COLUMNS = [c for c in EXPECTED_COLUMNS if c not in ("승인번호", "전자결재진행상태")]


@dataclass
class ReviewState:
    전표키: str
    검토상태: str = "미검토"
    메모: str = ""
    검토자: str = ""
    검토일시: str = ""
    내용해시: str = ""
    변경감지: bool = False
    이전검토상태: str = ""


def _cell(v) -> str:
    if v is None:
        return ""
    if isinstance(v, Decimal):
        return format(v.normalize(), "f")
    return str(v)


def content_hashes(df: pd.DataFrame) -> dict[str, str]:
    """전표키별 분개 내용 해시 (행 순서 무관)."""
    cols = [c for c in HASH_COLUMNS if c in df.columns]
    lines: dict[str, list[str]] = {}
    for key, *vals in df[["전표키", *cols]].itertuples(index=False, name=None):
        lines.setdefault(key, []).append("\x1f".join(_cell(v) for v in vals))
    return {k: hashlib.sha256("\x1e".join(sorted(v)).encode("utf-8")).hexdigest()[:16] for k, v in lines.items()}


def init_reviews(hashes: dict[str, str], previous: dict[str, ReviewState] | None = None) -> dict[str, ReviewState]:
    """현재 전표에 대한 검토 상태. 이전 저장분은 내용이 같을 때만 그대로 승계."""
    previous = previous or {}
    out = {}
    for key, h in hashes.items():
        prev = previous.get(key)
        if prev is None:
            out[key] = ReviewState(key, 내용해시=h)
        elif prev.내용해시 == h:
            out[key] = ReviewState(**{**asdict(prev), "변경감지": prev.변경감지})
        else:
            out[key] = ReviewState(key, "미검토", prev.메모, prev.검토자, prev.검토일시, h, True, prev.검토상태)
    return out


def update_review(reviews: dict[str, ReviewState], key: str, *, status: str | None = None, memo: str | None = None,
                  reviewer: str | None = None) -> ReviewState:
    r = reviews[key]
    if status is not None:
        if status not in REVIEW_STATES:
            raise ValueError(f"알 수 없는 검토상태: {status}")
        r.검토상태 = status
        if status == "검토완료":
            r.변경감지 = False  # 사용자가 변경 내용을 다시 검토 완료
    if memo is not None:
        r.메모 = memo
    if reviewer is not None:
        r.검토자 = reviewer
    r.검토일시 = dt.datetime.now().isoformat(timespec="seconds")
    return r


def reviews_frame(reviews: dict[str, ReviewState]) -> pd.DataFrame:
    return pd.DataFrame([asdict(r) for r in reviews.values()], columns=list(ReviewState.__dataclass_fields__))


# ------------------------------------------------------------------ 저장/재열기 ----
def _enc(v):
    if isinstance(v, Decimal):
        return {"__d": str(v)}
    if isinstance(v, (dt.date, dt.datetime)):
        return {"__t": v.isoformat()}
    return v


def _dec(v):
    if isinstance(v, dict) and "__d" in v:
        return Decimal(v["__d"])
    if isinstance(v, dict) and "__t" in v:
        return dt.date.fromisoformat(v["__t"][:10])
    return v


def safe_save_name(name: str) -> str:
    name = re.sub(r"[^\w가-힣\-. ]", "_", name).strip() or "검토저장"
    return name[:80]


def save_session(path: str | Path, *, source_name: str, source_sha256: str, sheet: str, header_row: int,
                 raw: pd.DataFrame, df: pd.DataFrame, reviews: dict[str, ReviewState], evidence_json: str,
                 settings: dict, preamble: list[str] | None = None) -> Path:
    """로컬(local_data/) 저장. 실데이터를 포함하므로 저장소에 올리지 않는다(.gitignore)."""
    payload = {
        "format": SAVE_FORMAT,
        "saved_at": dt.datetime.now().isoformat(timespec="seconds"),
        "source": {"name": source_name, "sha256": source_sha256, "sheet": sheet, "header_row": header_row,
                   "preamble": preamble or []},
        "raw": raw.to_dict("records"),
        "rows": [{k: _enc(v) for k, v in r.items()} for r in df.to_dict("records")],
        "reviews": [asdict(r) for r in reviews.values()],
        "evidence": json.loads(evidence_json),
        "settings": settings,
    }
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, default=str), encoding="utf-8")
    return path


def load_session(path: str | Path) -> dict:
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    if data.get("format") != SAVE_FORMAT:
        raise ValueError("지원하지 않는 저장 파일 형식입니다.")
    data["df"] = pd.DataFrame([{k: _dec(v) for k, v in r.items()} for r in data["rows"]], dtype=object)
    data["raw_df"] = pd.DataFrame(data["raw"], dtype=object)
    data["review_states"] = {r["전표키"]: ReviewState(**r) for r in data["reviews"]}
    return data


def list_saves(folder: str | Path) -> list[Path]:
    folder = Path(folder)
    return sorted(folder.glob("*.json"), key=lambda p: p.stat().st_mtime, reverse=True) if folder.exists() else []


def delete_save(path: str | Path, folder: str | Path) -> None:
    """저장 폴더 안의 저장 파일만 삭제 (원본 엑셀은 절대 삭제하지 않음)."""
    path, folder = Path(path).resolve(), Path(folder).resolve()
    if folder not in path.parents or path.suffix != ".json":
        raise ValueError("저장 폴더의 저장 파일만 삭제할 수 있습니다.")
    path.unlink()
