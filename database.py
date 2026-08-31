from __future__ import annotations

import os
import shutil
import sqlite3
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

from security import DataProtector


APP_DIR_NAME = "업무기한관리"
RELEASE_APP_DIR_NAME = "내일정관리"


def default_data_dir() -> Path:
    override = os.environ.get("DDAY_MANAGER_DATA_DIR")
    if override:
        return Path(override)
    base = Path(os.environ.get("LOCALAPPDATA", Path.home()))
    app_dir_name = RELEASE_APP_DIR_NAME if getattr(sys, "frozen", False) else APP_DIR_NAME
    return base / app_dir_name


class Database:
    def __init__(self, path: Path | None = None):
        self.data_dir = (path.parent if path else default_data_dir())
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.path = path or (self.data_dir / "schedule.db")
        self.conn = sqlite3.connect(self.path)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys = ON")
        self.conn.execute("PRAGMA journal_mode = WAL")
        self.protector: DataProtector | None = None
        self._create_schema()

    def _create_schema(self) -> None:
        self.conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS complaints (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                receipt_no TEXT NOT NULL UNIQUE,
                complaint_name TEXT NOT NULL,
                processing_stage TEXT NOT NULL DEFAULT '',
                received_at TEXT NOT NULL,
                registered_at TEXT NOT NULL DEFAULT '',
                original_deadline TEXT NOT NULL,
                current_deadline TEXT NOT NULL,
                applicant_name_enc TEXT NOT NULL DEFAULT '',
                birth_date_enc TEXT NOT NULL DEFAULT '',
                intake_type TEXT NOT NULL DEFAULT '',
                assignee_enc TEXT NOT NULL DEFAULT '',
                memo_enc TEXT NOT NULL DEFAULT '',
                status TEXT NOT NULL DEFAULT '진행',
                completed_at TEXT NOT NULL DEFAULT '',
                source_file TEXT NOT NULL DEFAULT '',
                progress_state TEXT NOT NULL DEFAULT '{}',
                stage_config TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS tasks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL,
                applicant_name_enc TEXT NOT NULL DEFAULT '',
                birth_date_enc TEXT NOT NULL DEFAULT '',
                registered_at TEXT NOT NULL,
                deadline TEXT NOT NULL,
                content_enc TEXT NOT NULL DEFAULT '',
                priority TEXT NOT NULL DEFAULT '',
                processing_stage TEXT NOT NULL DEFAULT '접수',
                progress_state TEXT NOT NULL DEFAULT '{}',
                stage_config TEXT NOT NULL DEFAULT '',
                status TEXT NOT NULL DEFAULT '진행',
                completed_at TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS calendar_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                event_date TEXT NOT NULL,
                event_time TEXT NOT NULL DEFAULT '',
                title TEXT NOT NULL,
                details_enc TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS deadline_extensions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                complaint_id INTEGER NOT NULL,
                previous_deadline TEXT NOT NULL,
                new_deadline TEXT NOT NULL,
                reason_enc TEXT NOT NULL DEFAULT '',
                extended_at TEXT NOT NULL,
                FOREIGN KEY (complaint_id) REFERENCES complaints(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS processing_stage_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                complaint_id INTEGER NOT NULL,
                previous_stage TEXT NOT NULL DEFAULT '',
                new_stage TEXT NOT NULL,
                changed_at TEXT NOT NULL,
                FOREIGN KEY (complaint_id) REFERENCES complaints(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS import_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                source_file TEXT NOT NULL,
                imported_at TEXT NOT NULL,
                new_count INTEGER NOT NULL,
                changed_count INTEGER NOT NULL,
                skipped_count INTEGER NOT NULL,
                error_count INTEGER NOT NULL
            );

            CREATE TABLE IF NOT EXISTS audit_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                entity_kind TEXT NOT NULL,
                entity_id INTEGER NOT NULL,
                action TEXT NOT NULL,
                details TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL
            );
            """
        )
        columns = {row["name"] for row in self.conn.execute("PRAGMA table_info(complaints)").fetchall()}
        if "progress_state" not in columns:
            self.conn.execute("ALTER TABLE complaints ADD COLUMN progress_state TEXT NOT NULL DEFAULT '{}'")
        if "stage_config" not in columns:
            self.conn.execute("ALTER TABLE complaints ADD COLUMN stage_config TEXT NOT NULL DEFAULT ''")
        task_columns = {row["name"] for row in self.conn.execute("PRAGMA table_info(tasks)").fetchall()}
        if "processing_stage" not in task_columns:
            self.conn.execute("ALTER TABLE tasks ADD COLUMN processing_stage TEXT NOT NULL DEFAULT '접수'")
        if "progress_state" not in task_columns:
            self.conn.execute("ALTER TABLE tasks ADD COLUMN progress_state TEXT NOT NULL DEFAULT '{}'")
        if "stage_config" not in task_columns:
            self.conn.execute("ALTER TABLE tasks ADD COLUMN stage_config TEXT NOT NULL DEFAULT ''")
        event_columns = {row["name"] for row in self.conn.execute("PRAGMA table_info(calendar_events)").fetchall()}
        if "event_time" not in event_columns:
            self.conn.execute("ALTER TABLE calendar_events ADD COLUMN event_time TEXT NOT NULL DEFAULT ''")
        self.conn.commit()

    def _log_audit(self, kind: str, item_id: int, action: str, details: str = "") -> None:
        self.conn.execute(
            "INSERT INTO audit_logs(entity_kind, entity_id, action, details, created_at) VALUES(?, ?, ?, ?, ?)",
            (kind, item_id, action, details, datetime.now().isoformat(timespec="seconds")),
        )

    def audit_history(self, kind: str, item_id: int) -> list[dict[str, Any]]:
        rows = self.conn.execute(
            "SELECT action, details, created_at FROM audit_logs WHERE entity_kind=? AND entity_id=? ORDER BY id DESC",
            (kind, item_id),
        ).fetchall()
        return [dict(row) for row in rows]

    def get_setting(self, key: str, default: str = "") -> str:
        row = self.conn.execute("SELECT value FROM settings WHERE key = ?", (key,)).fetchone()
        return row["value"] if row else default

    def set_setting(self, key: str, value: str) -> None:
        self.conn.execute(
            "INSERT INTO settings(key, value) VALUES(?, ?) "
            "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
            (key, value),
        )
        self.conn.commit()

    def has_password(self) -> bool:
        return bool(self.get_setting("password_verifier"))

    def set_initial_password(self, password: str) -> None:
        salt = DataProtector.new_salt()
        protector = DataProtector(password, salt)
        self.set_setting("password_salt", salt)
        self.set_setting("password_verifier", protector.make_verifier())
        self.protector = protector

    def enable_auto_unlock(self, password: str) -> None:
        from security import protect_for_windows_user

        self.set_setting("auto_unlock_secret", protect_for_windows_user(password))

    def try_auto_unlock(self) -> bool:
        from security import unprotect_for_windows_user

        secret = self.get_setting("auto_unlock_secret")
        if not secret:
            return False
        password = unprotect_for_windows_user(secret)
        self.unlock(password)
        return True

    def unlock(self, password: str) -> None:
        protector = DataProtector(password, self.get_setting("password_salt"))
        protector.verify(self.get_setting("password_verifier"))
        self.protector = protector

    def _enc(self, value: str | None) -> str:
        if not self.protector:
            raise RuntimeError("데이터베이스 잠금이 해제되지 않았습니다.")
        return self.protector.encrypt(value)

    def _dec(self, value: str | None) -> str:
        if not self.protector:
            raise RuntimeError("데이터베이스 잠금이 해제되지 않았습니다.")
        return self.protector.decrypt(value)

    def add_complaint(self, data: dict[str, str]) -> int:
        now = datetime.now().isoformat(timespec="seconds")
        try:
            cur = self.conn.execute(
                """
                INSERT INTO complaints (
                    receipt_no, complaint_name, processing_stage, received_at,
                    registered_at, original_deadline, current_deadline,
                    applicant_name_enc, birth_date_enc, intake_type, assignee_enc,
                    memo_enc, source_file, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    data["receipt_no"], data["complaint_name"], data.get("processing_stage", ""),
                    data["received_at"], data.get("registered_at", ""), data["deadline"],
                    data["deadline"], self._enc(data.get("applicant_name")),
                    self._enc(data.get("birth_date")), data.get("intake_type", ""),
                    self._enc(data.get("assignee")), self._enc(data.get("memo")),
                    data.get("source_file", ""), now, now,
                ),
            )
        except sqlite3.IntegrityError as exc:
            if "complaints.receipt_no" in str(exc):
                raise ValueError(
                    f"접수번호 '{data['receipt_no']}'는 이미 등록되어 있습니다.\n"
                    "기존 민원을 확인하거나 다른 접수번호를 입력하세요."
                ) from exc
            raise
        item_id = int(cur.lastrowid)
        self._log_audit("complaint", item_id, "민원 등록")
        self.conn.commit()
        return item_id

    def update_complaint_from_import(self, complaint_id: int, data: dict[str, str]) -> None:
        current = self.conn.execute(
            "SELECT current_deadline FROM complaints WHERE id = ?", (complaint_id,)
        ).fetchone()
        if not current:
            return
        if current["current_deadline"] != data["deadline"]:
            self.extend_deadline(
                complaint_id, data["deadline"], "엑셀 재가져오기에서 처리기한 변경 확인"
            )
        # For an existing receipt number, the imported workbook is authoritative
        # only for the current deadline. Preserve every other previously entered field.

    def complaint_by_receipt_no(self, receipt_no: str) -> sqlite3.Row | None:
        return self.conn.execute(
            "SELECT * FROM complaints WHERE receipt_no = ?", (receipt_no,)
        ).fetchone()

    def complaint_by_id(self, complaint_id: int) -> dict[str, Any] | None:
        row = self.conn.execute("SELECT * FROM complaints WHERE id = ?", (complaint_id,)).fetchone()
        return self._decode_complaint(row) if row else None

    def update_complaint(self, complaint_id: int, data: dict[str, str]) -> None:
        current = self.conn.execute("SELECT * FROM complaints WHERE id = ?", (complaint_id,)).fetchone()
        if not current:
            raise ValueError("민원 정보를 찾을 수 없습니다.")
        duplicate = self.conn.execute(
            "SELECT id FROM complaints WHERE receipt_no = ? AND id != ?",
            (data["receipt_no"], complaint_id),
        ).fetchone()
        if duplicate:
            raise ValueError(f"접수번호 '{data['receipt_no']}'는 이미 등록되어 있습니다.")
        now = datetime.now().isoformat(timespec="seconds")
        self.conn.execute(
            """
            UPDATE complaints SET receipt_no=?, complaint_name=?, received_at=?, registered_at=?,
                applicant_name_enc=?, birth_date_enc=?, intake_type=?, assignee_enc=?, memo_enc=?, updated_at=?
            WHERE id=?
            """,
            (
                data["receipt_no"], data["complaint_name"], data["received_at"],
                data.get("registered_at", ""), self._enc(data.get("applicant_name")),
                self._enc(data.get("birth_date")), data.get("intake_type", ""),
                self._enc(data.get("assignee")), self._enc(data.get("memo")), now, complaint_id,
            ),
        )
        self._log_audit("complaint", complaint_id, "민원 기본정보 수정", "접수번호·민원명·민원인·접수일자·메모 등")
        self.conn.commit()
        if current["current_deadline"] != data["deadline"]:
            self.extend_deadline(complaint_id, data["deadline"], "민원 편집에서 처리기한 변경")
        if current["processing_stage"] != data["processing_stage"]:
            self.update_processing_stage(complaint_id, data["processing_stage"])

    def update_processing_stage(self, complaint_id: int, new_stage: str) -> None:
        row = self.conn.execute(
            "SELECT processing_stage FROM complaints WHERE id = ?", (complaint_id,)
        ).fetchone()
        if not row:
            raise ValueError("민원 정보를 찾을 수 없습니다.")
        previous = row["processing_stage"]
        if previous == new_stage:
            return
        now = datetime.now().isoformat(timespec="seconds")
        self.conn.execute(
            "INSERT INTO processing_stage_history(complaint_id, previous_stage, new_stage, changed_at) VALUES (?, ?, ?, ?)",
            (complaint_id, previous, new_stage, now),
        )
        self.conn.execute(
            "UPDATE complaints SET processing_stage = ?, updated_at = ? WHERE id = ?",
            (new_stage, now, complaint_id),
        )
        self._log_audit("complaint", complaint_id, "진행 단계 변경", f"{previous} → {new_stage}")
        self.conn.commit()

    def processing_stage_history(self, complaint_id: int) -> list[dict[str, str]]:
        rows = self.conn.execute(
            "SELECT previous_stage, new_stage, changed_at FROM processing_stage_history WHERE complaint_id = ? ORDER BY id DESC",
            (complaint_id,),
        ).fetchall()
        return [dict(row) for row in rows]

    def update_progress_state(self, complaint_id: int, state: dict[str, Any], processing_stage: str) -> None:
        import json

        if not self.conn.execute("SELECT id FROM complaints WHERE id = ?", (complaint_id,)).fetchone():
            raise ValueError("민원 정보를 찾을 수 없습니다.")
        self.conn.execute(
            "UPDATE complaints SET progress_state = ?, updated_at = ? WHERE id = ?",
            (json.dumps(state, ensure_ascii=False), datetime.now().isoformat(timespec="seconds"), complaint_id),
        )
        self._log_audit("complaint", complaint_id, "진행 현황 수정")
        self.conn.commit()
        self.update_processing_stage(complaint_id, processing_stage)

    def extend_deadline(self, complaint_id: int, new_deadline: str, reason: str) -> None:
        row = self.conn.execute(
            "SELECT current_deadline FROM complaints WHERE id = ?", (complaint_id,)
        ).fetchone()
        if not row:
            raise ValueError("민원 정보를 찾을 수 없습니다.")
        now = datetime.now().isoformat(timespec="seconds")
        self.conn.execute(
            "INSERT INTO deadline_extensions(complaint_id, previous_deadline, new_deadline, reason_enc, extended_at) VALUES (?, ?, ?, ?, ?)",
            (complaint_id, row["current_deadline"], new_deadline, self._enc(reason), now),
        )
        self.conn.execute(
            "UPDATE complaints SET current_deadline=?, updated_at=? WHERE id=?",
            (new_deadline, now, complaint_id),
        )
        self._log_audit("complaint", complaint_id, "처리기한 변경", f"{row['current_deadline']} → {new_deadline} · {reason}")
        self.conn.commit()

    def add_task(self, data: dict[str, str]) -> int:
        now = datetime.now().isoformat(timespec="seconds")
        cur = self.conn.execute(
            """
            INSERT INTO tasks(title, applicant_name_enc, birth_date_enc, registered_at,
                deadline, content_enc, priority, processing_stage, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                data["title"], self._enc(data.get("applicant_name")),
                self._enc(data.get("birth_date")), data["registered_at"], data["deadline"],
                self._enc(data.get("content")), data.get("priority", ""),
                data.get("processing_stage", "접수"), now, now,
            ),
        )
        item_id = int(cur.lastrowid)
        self._log_audit("task", item_id, "수시업무 등록")
        self.conn.commit()
        return item_id

    def task_by_id(self, task_id: int) -> dict[str, Any] | None:
        row = self.conn.execute("SELECT * FROM tasks WHERE id = ?", (task_id,)).fetchone()
        return self._decode_task(row) if row else None

    def update_task(self, task_id: int, data: dict[str, str]) -> None:
        if not self.conn.execute("SELECT id FROM tasks WHERE id = ?", (task_id,)).fetchone():
            raise ValueError("수시업무를 찾을 수 없습니다.")
        now = datetime.now().isoformat(timespec="seconds")
        self.conn.execute(
            """
            UPDATE tasks SET title=?, registered_at=?, deadline=?, content_enc=?, priority=?, updated_at=?
            WHERE id=?
            """,
            (
                data["title"], data["registered_at"], data["deadline"],
                self._enc(data.get("content")), data.get("priority", ""), now, task_id,
            ),
        )
        self._log_audit("task", task_id, "수시업무 기본정보 수정", "업무 제목·등록일·처리기한·비고·내용")
        self.conn.commit()

    def update_task_progress_state(self, task_id: int, state: dict[str, Any], processing_stage: str) -> None:
        import json

        if not self.conn.execute("SELECT id FROM tasks WHERE id = ?", (task_id,)).fetchone():
            raise ValueError("수시업무를 찾을 수 없습니다.")
        self.conn.execute(
            "UPDATE tasks SET progress_state = ?, processing_stage = ?, updated_at = ? WHERE id = ?",
            (json.dumps(state, ensure_ascii=False), processing_stage, datetime.now().isoformat(timespec="seconds"), task_id),
        )
        self._log_audit("task", task_id, "진행 현황 수정", processing_stage)
        self.conn.commit()

    def update_item_stage_config(self, kind: str, item_id: int, config: list[dict[str, Any]]) -> None:
        import json

        table = "complaints" if kind == "complaint" else "tasks"
        self.conn.execute(
            f"UPDATE {table} SET stage_config=?, updated_at=? WHERE id=?",  # noqa: S608
            (json.dumps(config, ensure_ascii=False), datetime.now().isoformat(timespec="seconds"), item_id),
        )
        self._log_audit(kind, item_id, "개별 진행 현황 정의 수정")
        self.conn.commit()

    def bulk_update_tasks(self, task_ids: list[int], deadline: str = "", priority: str = "") -> None:
        now = datetime.now().isoformat(timespec="seconds")
        for task_id in task_ids:
            if deadline:
                self.conn.execute("UPDATE tasks SET deadline=?, updated_at=? WHERE id=?", (deadline, now, task_id))
            if priority:
                self.conn.execute("UPDATE tasks SET priority=?, updated_at=? WHERE id=?", (priority, now, task_id))
            changes = " · ".join(part for part in (f"처리기한 {deadline}" if deadline else "", f"비고 {priority}" if priority else "") if part)
            if changes:
                self._log_audit("task", task_id, "일괄 변경", changes)
        self.conn.commit()

    def list_complaints(self, include_completed: bool = False) -> list[dict[str, Any]]:
        where = "" if include_completed else "WHERE status = '진행'"
        rows = self.conn.execute(
            f"SELECT * FROM complaints {where} ORDER BY current_deadline, id"  # noqa: S608
        ).fetchall()
        return [self._decode_complaint(row) for row in rows]

    def list_tasks(self, include_completed: bool = False) -> list[dict[str, Any]]:
        where = "" if include_completed else "WHERE status = '진행'"
        rows = self.conn.execute(
            f"SELECT * FROM tasks {where} ORDER BY deadline, id"  # noqa: S608
        ).fetchall()
        return [self._decode_task(row) for row in rows]

    def _decode_complaint(self, row: sqlite3.Row) -> dict[str, Any]:
        item = dict(row)
        item["applicant_name"] = self._dec(item.pop("applicant_name_enc"))
        item["birth_date"] = self._dec(item.pop("birth_date_enc"))
        item["assignee"] = self._dec(item.pop("assignee_enc"))
        item["memo"] = self._dec(item.pop("memo_enc"))
        return item

    def _decode_task(self, row: sqlite3.Row) -> dict[str, Any]:
        item = dict(row)
        item["applicant_name"] = self._dec(item.pop("applicant_name_enc"))
        item["birth_date"] = self._dec(item.pop("birth_date_enc"))
        item["content"] = self._dec(item.pop("content_enc"))
        return item

    def mark_complete(self, kind: str, item_id: int) -> None:
        table = "complaints" if kind == "complaint" else "tasks"
        now = datetime.now().isoformat(timespec="seconds")
        self.conn.execute(
            f"UPDATE {table} SET status='완료', completed_at=?, updated_at=? WHERE id=?",  # noqa: S608
            (now, now, item_id),
        )
        self._log_audit(kind, item_id, "처리 완료")
        completed = self.conn.execute(
            """
            SELECT kind, id FROM (
                SELECT 'complaint' AS kind, id, completed_at FROM complaints WHERE status='완료'
                UNION ALL
                SELECT 'task' AS kind, id, completed_at FROM tasks WHERE status='완료'
            )
            ORDER BY completed_at DESC, kind DESC, id DESC
            """
        ).fetchall()
        for old in completed[30:]:
            old_table = "complaints" if old["kind"] == "complaint" else "tasks"
            self.conn.execute(f"DELETE FROM {old_table} WHERE id=?", (old["id"],))  # noqa: S608
        self.conn.commit()

    def restore_completed(self, kind: str, item_id: int) -> None:
        table = "complaints" if kind == "complaint" else "tasks"
        now = datetime.now().isoformat(timespec="seconds")
        self.conn.execute(
            f"UPDATE {table} SET status='진행', completed_at='', updated_at=? WHERE id=?",  # noqa: S608
            (now, item_id),
        )
        self._log_audit(kind, item_id, "완료 내역에서 복원")
        self.conn.commit()

    def move_to_trash(self, kind: str, item_ids: list[int]) -> None:
        if not item_ids:
            return
        table = "complaints" if kind == "complaint" else "tasks"
        now = datetime.now().isoformat(timespec="seconds")
        placeholders = ",".join("?" for _ in item_ids)
        self.conn.execute(
            f"UPDATE {table} SET status='삭제', updated_at=? WHERE id IN ({placeholders})",  # noqa: S608
            (now, *item_ids),
        )
        for item_id in item_ids:
            self._log_audit(kind, item_id, "휴지통으로 이동")
        self.conn.commit()

    def restore_from_trash(self, kind: str, item_ids: list[int]) -> None:
        if not item_ids:
            return
        table = "complaints" if kind == "complaint" else "tasks"
        now = datetime.now().isoformat(timespec="seconds")
        placeholders = ",".join("?" for _ in item_ids)
        self.conn.execute(
            f"UPDATE {table} SET status='진행', completed_at='', updated_at=? WHERE id IN ({placeholders})",  # noqa: S608
            (now, *item_ids),
        )
        for item_id in item_ids:
            self._log_audit(kind, item_id, "휴지통에서 복원")
        self.conn.commit()

    def permanently_delete(self, kind: str, item_ids: list[int]) -> None:
        if not item_ids:
            return
        table = "complaints" if kind == "complaint" else "tasks"
        placeholders = ",".join("?" for _ in item_ids)
        for item_id in item_ids:
            self._log_audit(kind, item_id, "영구 삭제")
        self.conn.execute(
            f"DELETE FROM {table} WHERE status='삭제' AND id IN ({placeholders})",  # noqa: S608
            tuple(item_ids),
        )
        self.conn.commit()

    def extension_history(self, complaint_id: int) -> list[dict[str, str]]:
        rows = self.conn.execute(
            "SELECT * FROM deadline_extensions WHERE complaint_id=? ORDER BY id", (complaint_id,)
        ).fetchall()
        result = []
        for row in rows:
            item = dict(row)
            item["reason"] = self._dec(item.pop("reason_enc"))
            result.append(item)
        return result

    def log_import(self, source_file: str, counts: dict[str, int]) -> None:
        self.conn.execute(
            "INSERT INTO import_logs(source_file, imported_at, new_count, changed_count, skipped_count, error_count) VALUES (?, ?, ?, ?, ?, ?)",
            (
                source_file, datetime.now().isoformat(timespec="seconds"), counts.get("new", 0),
                counts.get("changed", 0), counts.get("skipped", 0), counts.get("error", 0),
            ),
        )
        self.conn.commit()

    def add_calendar_event(self, event_date: str, title: str, details: str = "", event_time: str = "") -> int:
        now = datetime.now().isoformat(timespec="seconds")
        cursor = self.conn.execute(
            "INSERT INTO calendar_events(event_date, event_time, title, details_enc, created_at, updated_at) VALUES(?, ?, ?, ?, ?, ?)",
            (event_date, event_time, title, self._enc(details), now, now),
        )
        event_id = int(cursor.lastrowid)
        self._log_audit("calendar_event", event_id, "일정 등록", event_date)
        self.conn.commit()
        return event_id

    def list_calendar_events(self) -> list[dict[str, Any]]:
        rows = self.conn.execute(
            "SELECT * FROM calendar_events ORDER BY event_date, id"
        ).fetchall()
        result: list[dict[str, Any]] = []
        for row in rows:
            item = dict(row)
            item["details"] = self._dec(item.pop("details_enc"))
            result.append(item)
        return result

    def calendar_event_by_id(self, event_id: int) -> dict[str, Any] | None:
        row = self.conn.execute("SELECT * FROM calendar_events WHERE id=?", (event_id,)).fetchone()
        if not row:
            return None
        item = dict(row)
        item["details"] = self._dec(item.pop("details_enc"))
        return item

    def update_calendar_event(
        self,
        event_id: int,
        event_date: str,
        title: str,
        details: str = "",
        event_time: str = "",
    ) -> None:
        if not self.conn.execute("SELECT id FROM calendar_events WHERE id=?", (event_id,)).fetchone():
            raise ValueError("수정할 일정을 찾을 수 없습니다.")
        now = datetime.now().isoformat(timespec="seconds")
        self.conn.execute(
            "UPDATE calendar_events SET event_date=?, event_time=?, title=?, details_enc=?, updated_at=? WHERE id=?",
            (event_date, event_time, title, self._enc(details), now, event_id),
        )
        self._log_audit("calendar_event", event_id, "일정 수정", f"{event_date} {event_time}".strip())
        self.conn.commit()

    def delete_calendar_event(self, event_id: int) -> None:
        if not self.conn.execute("SELECT id FROM calendar_events WHERE id=?", (event_id,)).fetchone():
            return
        self._log_audit("calendar_event", event_id, "일정 삭제")
        self.conn.execute("DELETE FROM calendar_events WHERE id=?", (event_id,))
        self.conn.commit()

    def backup(self) -> Path:
        backup_dir = self.data_dir / "backups"
        backup_dir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        target = backup_dir / f"schedule_{stamp}.db"
        dst = sqlite3.connect(target)
        try:
            self.conn.backup(dst)
        finally:
            dst.close()
        backups = sorted(backup_dir.glob("schedule_*.db"), reverse=True)
        for old in backups[30:]:
            old.unlink(missing_ok=True)
        return target

    def close(self) -> None:
        self.conn.close()
