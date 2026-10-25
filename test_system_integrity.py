"""
Automated Test Suite for Aircraft Battery Maintenance Monitoring System
Tests data integrity, schema validation, atomic routing, and date utilities.
"""

import os
import shutil
import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path
import pandas as pd

from config import load_app_config, get_master_excel_path
from validation import (
    validate_workbook_schema,
    validate_battery_unit_record,
    validate_fsb_record,
    validate_scrap_record,
    parse_date_flexible,
    format_date_with_month_name,
    to_date_object
)
from backup_manager import create_backup, list_backups, get_backup_directory
from sync_service import get_file_fingerprint, check_external_conflict
from excel_repository import ExcelRepository


class TestSystemIntegrity(unittest.TestCase):

    def setUp(self):
        self.test_dir = Path(tempfile.mkdtemp(prefix="battery_test_"))
        self.master_src = Path("Daily Battery.xlsx").resolve()
        self.test_master = self.test_dir / "Daily Battery.xlsx"
        if self.master_src.exists():
            shutil.copy2(str(self.master_src), str(self.test_master))

    def tearDown(self):
        if self.test_dir.exists():
            shutil.rmtree(str(self.test_dir), ignore_errors=True)

    def test_date_parsing_and_formatting(self):
        """Tests that Indonesian month strings and standard dates parse correctly."""
        # Standard ISO
        d1 = parse_date_flexible("2026-10-20")
        self.assertEqual(d1, date(2026, 10, 20))

        # Indonesian long month
        d2 = parse_date_flexible("15 Oktober 2026")
        self.assertEqual(d2, date(2026, 10, 15))

        # Indonesian short month
        d3 = parse_date_flexible("14 Okt 2026")
        self.assertEqual(d3, date(2026, 10, 14))

        # Edge cases: None, empty, '-'
        self.assertIsNone(parse_date_flexible(None))
        self.assertIsNone(parse_date_flexible("-"))
        self.assertIsNone(parse_date_flexible(""))

        # Formatting
        fmt = format_date_with_month_name(date(2026, 10, 20))
        self.assertEqual(fmt, "20 Oktober 2026")

        # Alias compatibility
        self.assertEqual(to_date_object("1 Januari 2025"), date(2025, 1, 1))

    def test_validation_rules(self):
        """Tests validation for required fields, TAT constraints, and duplicate WO."""
        # Valid record
        valid_rec = {
            "WO": "WO-9999",
            "Part Number": "1615-1",
            "Serial Number": "SN-001",
            "Customer": "Garuda Indonesia",
            "Status": "Overhaul",
            "TAT": 5
        }
        ok, msg = validate_battery_unit_record(valid_rec, existing_wos={"WO-0001", "WO-0002"})
        self.assertTrue(ok)
        self.assertEqual(msg, "")

        # Duplicate WO
        dup_rec = dict(valid_rec, WO="WO-0001")
        ok_dup, msg_dup = validate_battery_unit_record(dup_rec, existing_wos={"WO-0001", "WO-0002"})
        self.assertFalse(ok_dup)
        self.assertIn("sudah terdaftar aktif", msg_dup)

        # Missing required field (WO)
        empty_wo_rec = dict(valid_rec, WO="")
        ok_empty, msg_empty = validate_battery_unit_record(empty_wo_rec)
        self.assertFalse(ok_empty)
        self.assertIn("Nomor WO", msg_empty)

        # Negative TAT
        neg_tat_rec = dict(valid_rec, TAT=-3)
        ok_tat, msg_tat = validate_battery_unit_record(neg_tat_rec)
        self.assertFalse(ok_tat)
        self.assertIn("tidak boleh negatif", msg_tat)

    def test_workbook_schema_validation(self):
        """Tests schema validation on the master excel file."""
        if not self.test_master.exists():
            self.skipTest("Master Excel not available in test environment.")
        is_valid, errors, warnings = validate_workbook_schema(self.test_master)
        self.assertTrue(is_valid, f"Workbook validation failed: {errors}")
        self.assertEqual(len(errors), 0)

    def test_excel_repository_read_and_atomic_route(self):
        """Tests ExcelRepository reading and atomic routing from Daily to History FSB."""
        if not self.test_master.exists():
            self.skipTest("Master Excel not available in test environment.")

        repo = ExcelRepository(self.test_master)
        data, warns = repo.load_all_data()
        self.assertIn("daily", data)
        self.assertIn("fsb", data)
        self.assertIn("scrap", data)

        initial_daily = data["daily"]
        if initial_daily.empty:
            self.skipTest("No daily records to test routing.")

        target_wo = str(initial_daily.iloc[0]["WO"]).strip()
        initial_daily_count = len(initial_daily)
        initial_fsb_count = len(data["fsb"])

        # Execute route_unit_to_fsb
        ok, msg = repo.route_unit_to_fsb(target_wo, completion_date=date.today())
        self.assertTrue(ok, f"Routing to FSB failed: {msg}")

        # Reload and verify
        reloaded, _ = repo.load_all_data()
        new_daily = reloaded["daily"]
        new_fsb = reloaded["fsb"]

        self.assertEqual(len(new_daily), initial_daily_count - 1)
        self.assertEqual(len(new_fsb), initial_fsb_count + 1)
        self.assertFalse((new_daily["WO"].astype(str).str.strip() == target_wo).any())
        self.assertTrue((new_fsb["WO"].astype(str).str.strip() == target_wo).any())

    def test_backup_and_fingerprint(self):
        """Tests snapshot creation and fingerprint conflict detection."""
        if not self.test_master.exists():
            self.skipTest("Master Excel not available in test environment.")

        fp_initial = get_file_fingerprint(self.test_master)
        self.assertTrue(fp_initial["exists"])
        self.assertGreater(fp_initial["size"], 0)

        # No conflict when nothing changed
        conflict, _ = check_external_conflict(self.test_master, fp_initial)
        self.assertFalse(conflict)

        # Create backup
        bk_path = create_backup(self.test_master, reason="unit_test")
        self.assertIsNotNone(bk_path)
        self.assertTrue(bk_path.exists())


if __name__ == "__main__":
    unittest.main()
