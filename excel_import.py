from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime, time
from pathlib import Path
import sys
from typing import Any

VENDOR_DIR = Path(__file__).resolve().parent / ".vendor"
if VENDOR_DIR.exists() and str(VENDOR_DIR) not in sys.path:
    sys.path.insert(0, str(VENDOR_DIR))

import pandas as pd

from database import Database


ALIASES = {
    "receipt_no": ["접수번호", "민원접수번호", "신청번호"],
    "complaint_name": ["민원명", "민원종류", "신청보장", "신청명"],
    "processing_stage": ["처리단계", "진행상태", "처리상태"],
    "received_at": ["받은일자", "받은날짜", "수신일자"],
    "registered_at": ["접수일자", "접수일", "등록일자"],
    "deadline": ["처리기한", "처리기한일", "기한"],
    "applicant_name": ["민원인", "민원인명", "성명", "대상자"],
    "intake_type": ["접수구분", "접수경로", "신청구분"],
    "assignee": ["담당자", "처리담당자"],
}


def normalize_header(value: Any) -> str:
    return "".join(str(value).strip().replace("\n", "").split()).lower()


def discover_columns(columns: list[Any]) -> dict[str, Any]:
    normalized = {normalize_header(col): col for col in columns}
    mapping: dict[str, Any] = {}
    for field_name, aliases in ALIASES.items():
        for alias in aliases:
            key = normalize_header(alias)
            if key in normalized:
                mapping[field_name] = normalized[key]
                break
    return mapping


def format_datetime(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, pd.Timestamp):
        value = value.to_pydatetime()
    if isinstance(value, datetime):
        if value.time() == time.min:
            return value.strftime("%Y-%m-%d")
        return value.strftime("%Y-%m-%d %H:%M")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return ""
    parsed = pd.to_datetime(text, errors="coerce")
    if pd.isna(parsed):
        return text
    if parsed.time() == time.min:
        return parsed.strftime("%Y-%m-%d")
    return parsed.strftime("%Y-%m-%d %H:%M")


def read_workbook(path: Path) -> tuple[pd.DataFrame, dict[str, Any]]:
    suffix = path.suffix.lower()
    if suffix == ".xls":
        try:
            frame = pd.read_excel(path, engine="xlrd")
        except ImportError as exc:
            raise RuntimeError("구형 .xls 파일을 읽으려면 xlrd 모듈이 필요합니다.") from exc
    elif suffix in {".xlsx", ".xlsm"}:
        frame = pd.read_excel(path, engine="openpyxl")
    else:
        raise ValueError(".xls, .xlsx 또는 .xlsm 파일만 가져올 수 있습니다.")
    frame = frame.dropna(how="all")
    mapping = discover_columns(list(frame.columns))
    required = {"receipt_no", "complaint_name", "received_at", "deadline"}
    missing = required - set(mapping)
    if missing:
        labels = {key: ALIASES[key][0] for key in missing}
        raise ValueError("필수 열을 찾지 못했습니다: " + ", ".join(labels.values()))
    return frame, mapping


@dataclass
class ImportItem:
    row_number: int
    status: str
    message: str
    data: dict[str, str] = field(default_factory=dict)
    existing_id: int | None = None


def preview_import(path: Path, db: Database) -> list[ImportItem]:
    frame, mapping = read_workbook(path)
    items: list[ImportItem] = []
    seen: set[str] = set()
    for index, row in frame.iterrows():
        data: dict[str, str] = {}
        for field_name, column in mapping.items():
            raw = row[column]
            if field_name in {"received_at", "registered_at", "deadline"}:
                data[field_name] = format_datetime(raw)
            else:
                data[field_name] = "" if pd.isna(raw) else str(raw).strip()
        data["source_file"] = path.name
        receipt_no = data.get("receipt_no", "")
        required_values = [receipt_no, data.get("complaint_name", ""), data.get("received_at", ""), data.get("deadline", "")]
        if not all(required_values):
            items.append(ImportItem(index + 2, "오류", "필수 값 누락", data))
            continue
        if receipt_no in seen:
            items.append(ImportItem(index + 2, "오류", "파일 안에서 접수번호 중복", data))
            continue
        seen.add(receipt_no)
        existing = db.complaint_by_receipt_no(receipt_no)
        if not existing:
            items.append(ImportItem(index + 2, "신규", "새 민원", data))
        else:
            changed = existing["current_deadline"] != data["deadline"]
            status = "변경" if changed else "중복"
            message = "처리기한만 갱신" if changed else "기존 등록 건과 동일"
            items.append(ImportItem(index + 2, status, message, data, int(existing["id"])))
    return items


def apply_import(path: Path, db: Database, items: list[ImportItem]) -> dict[str, int]:
    counts = {"new": 0, "changed": 0, "skipped": 0, "error": 0}
    for item in items:
        if item.status == "신규":
            db.add_complaint(item.data)
            counts["new"] += 1
        elif item.status == "변경" and item.existing_id:
            db.update_complaint_from_import(item.existing_id, item.data)
            counts["changed"] += 1
        elif item.status == "오류":
            counts["error"] += 1
        else:
            counts["skipped"] += 1
    db.log_import(path.name, counts)
    return counts
