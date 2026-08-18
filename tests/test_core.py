from __future__ import annotations

import sqlite3
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from database import Database
from excel_import import apply_import, preview_import


class CoreFlowTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.db = Database(self.root / "schedule.db")
        self.db.set_initial_password("test-password-123")

    def tearDown(self) -> None:
        self.db.close()
        self.temp.cleanup()

    def _write_excel(
        self,
        deadline: str = "2026-09-09 18:00",
        stage: str = "접수",
        received_at: str = "2026-08-10",
        applicant_name: str = "가상민원인",
    ) -> Path:
        path = self.root / "받은민원.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.append(["순번", "접수번호", "민원명", "처리단계", "받은일자", "접수일자", "처리기한", "민원인", "접수구분"])
        ws.append([1, "R-2026-001", "맞춤형복지급여", stage, received_at, "2026-08-11", deadline, applicant_name, "연계"])
        wb.save(path)
        return path

    def test_import_duplicate_and_deadline_change_history(self) -> None:
        path = self._write_excel()
        first = preview_import(path, self.db)
        self.assertEqual([item.status for item in first], ["신규"])
        counts = apply_import(path, self.db, first)
        self.assertEqual(counts["new"], 1)

        duplicate = preview_import(path, self.db)
        self.assertEqual([item.status for item in duplicate], ["중복"])

        self._write_excel(
            "2026-10-09 18:00",
            stage="처리완료",
            received_at="2026-08-30",
            applicant_name="엑셀에서변경된이름",
        )
        changed = preview_import(path, self.db)
        self.assertEqual([item.status for item in changed], ["변경"])
        apply_import(path, self.db, changed)
        complaint = self.db.list_complaints()[0]
        self.assertEqual(complaint["original_deadline"], "2026-09-09 18:00")
        self.assertEqual(complaint["current_deadline"], "2026-10-09 18:00")
        self.assertEqual(complaint["processing_stage"], "접수")
        self.assertEqual(complaint["received_at"], "2026-08-10")
        self.assertEqual(complaint["applicant_name"], "가상민원인")
        self.assertEqual(len(self.db.extension_history(complaint["id"])), 1)

    def test_sensitive_values_are_not_plaintext_in_database(self) -> None:
        self.db.add_task(
            {
                "title": "서류 확인",
                "applicant_name": "가상민원인",
                "birth_date": "1950-03-12",
                "registered_at": "2026-08-15",
                "deadline": "2026-08-20",
                "content": "가상의 민감 메모",
                "priority": "높음",
            }
        )
        raw_conn = sqlite3.connect(self.db.path)
        try:
            raw = raw_conn.execute(
                "SELECT applicant_name_enc, birth_date_enc, content_enc FROM tasks"
            ).fetchone()
        finally:
            raw_conn.close()
        joined = " ".join(raw)
        self.assertNotIn("가상민원인", joined)
        self.assertNotIn("1950-03-12", joined)
        self.assertNotIn("가상의 민감 메모", joined)
        self.assertEqual(self.db.list_tasks()[0]["applicant_name"], "가상민원인")

    def test_backup_is_created(self) -> None:
        backup = self.db.backup()
        self.assertTrue(backup.exists())
        self.assertGreater(backup.stat().st_size, 0)

    def test_processing_stage_change_is_recorded(self) -> None:
        complaint_id = self.db.add_complaint(
            {
                "receipt_no": "R-STAGE-001",
                "complaint_name": "단계 기록 시험",
                "processing_stage": "1. 공적자료요청",
                "received_at": "2026-08-15",
                "registered_at": "2026-08-15",
                "deadline": "2026-09-15",
                "applicant_name": "가상민원인",
            }
        )
        self.db.update_processing_stage(complaint_id, "3. 조사중 > 3-2. 가정방문")
        complaint = self.db.complaint_by_id(complaint_id)
        self.assertEqual(complaint["processing_stage"], "3. 조사중 > 3-2. 가정방문")
        history = self.db.processing_stage_history(complaint_id)
        self.assertEqual(history[0]["previous_stage"], "1. 공적자료요청")
        self.assertEqual(history[0]["new_stage"], "3. 조사중 > 3-2. 가정방문")

    def test_duplicate_receipt_number_has_friendly_message(self) -> None:
        data = {
            "receipt_no": "R-DUPLICATE-001",
            "complaint_name": "중복 확인",
            "processing_stage": "1. 공적자료요청",
            "received_at": "2026-08-15",
            "deadline": "2026-09-15",
        }
        self.db.add_complaint(data)
        with self.assertRaisesRegex(ValueError, "이미 등록되어 있습니다"):
            self.db.add_complaint(data)

    def test_auto_unlock_uses_current_windows_user(self) -> None:
        self.db.enable_auto_unlock("test-password-123")
        self.db.protector = None
        self.assertTrue(self.db.try_auto_unlock())
        self.assertIsNotNone(self.db.protector)

    def test_complaint_and_task_can_be_fully_edited(self) -> None:
        complaint_id = self.db.add_complaint(
            {
                "receipt_no": "R-EDIT-OLD",
                "complaint_name": "수정 전 민원",
                "processing_stage": "1. 공적자료요청",
                "received_at": "2026-08-15",
                "registered_at": "2026-08-15",
                "deadline": "2026-09-15",
                "applicant_name": "수정 전 이름",
            }
        )
        self.db.update_complaint(
            complaint_id,
            {
                "receipt_no": "R-EDIT-NEW",
                "complaint_name": "수정 후 민원",
                "processing_stage": "3. 조사중 > 3-1. 소득조사",
                "received_at": "2026-08-16",
                "registered_at": "2026-08-16",
                "deadline": "2026-10-15",
                "applicant_name": "수정 후 이름",
                "birth_date": "1991-11-01",
                "intake_type": "수기",
                "assignee": "담당자",
                "memo": "수정 메모",
            },
        )
        complaint = self.db.complaint_by_id(complaint_id)
        self.assertEqual(complaint["receipt_no"], "R-EDIT-NEW")
        self.assertEqual(complaint["applicant_name"], "수정 후 이름")
        self.assertEqual(complaint["current_deadline"], "2026-10-15")
        self.assertEqual(complaint["original_deadline"], "2026-09-15")

        task_id = self.db.add_task(
            {"title": "수정 전", "registered_at": "2026-08-15", "deadline": "2026-08-20", "content": "전"}
        )
        self.db.update_task(
            task_id,
            {"title": "수정 후", "registered_at": "2026-08-16", "deadline": "2026-08-25", "priority": "긴급", "content": "후"},
        )
        task = self.db.task_by_id(task_id)
        self.assertEqual(task["title"], "수정 후")
        self.assertEqual(task["priority"], "긴급")

    def test_progress_state_is_saved_with_current_and_completed_nodes(self) -> None:
        complaint_id = self.db.add_complaint(
            {
                "receipt_no": "R-PROGRESS-001",
                "complaint_name": "진행 현황 시험",
                "processing_stage": "1. 신청",
                "received_at": "2026-08-15",
                "deadline": "2026-09-15",
            }
        )
        state = {
            "current": "b:investigation:income-check",
            "completed": ["s:application", "s:public-request"],
        }
        self.db.update_progress_state(
            complaint_id,
            state,
            "3. 조사중 > 3-1. 소득조사",
        )
        complaint = self.db.complaint_by_id(complaint_id)
        self.assertEqual(complaint["processing_stage"], "3. 조사중 > 3-1. 소득조사")
        self.assertIn("s:application", complaint["progress_state"])

    def test_completed_items_can_be_restored(self) -> None:
        complaint_id = self.db.add_complaint(
            {"receipt_no": "R-RESTORE", "complaint_name": "복원 시험", "received_at": "2026-08-15", "deadline": "2026-09-15"}
        )
        self.db.mark_complete("complaint", complaint_id)
        self.assertEqual(self.db.complaint_by_id(complaint_id)["status"], "완료")
        self.db.restore_completed("complaint", complaint_id)
        restored = self.db.complaint_by_id(complaint_id)
        self.assertEqual(restored["status"], "진행")
        self.assertEqual(restored["completed_at"], "")

    def test_task_progress_stages_are_saved(self) -> None:
        task_id = self.db.add_task(
            {"title": "수시 단계 시험", "registered_at": "2026-08-15", "deadline": "2026-09-15", "processing_stage": "접수"}
        )
        self.db.update_task_progress_state(
            task_id,
            {"current": "s:task-working", "completed": ["s:task-received"]},
            "작업중",
        )
        task = self.db.task_by_id(task_id)
        self.assertEqual(task["processing_stage"], "작업중")
        self.assertIn("task-received", task["progress_state"])


if __name__ == "__main__":
    unittest.main()
