"""
Excel Repository for Aircraft Battery Maintenance Monitoring System
Handles safe, atomic read and write operations, worksheet preservation,
table boundary management, formatting retention, formula preservation,
and robust error resilience across all three modules:
1. Daily Battery -> Worksheet 'Dashboard' (Table: DASHBOARD)
2. History FSB   -> Worksheet 'History FSB' (Table: Table2)
3. History SCRAP -> Worksheet 'History SCRAP' (Table: SCRAP)
"""

import os
import re
import json
import shutil
from copy import copy
from datetime import datetime, date
from pathlib import Path
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter, coordinate_to_tuple

from config import BASE_DIR
from backup_manager import create_backup
from sync_service import is_file_locked
from validation import (
    validate_workbook_schema,
    EXPECTED_DAILY_COLUMNS,
    EXPECTED_FSB_COLUMNS,
    EXPECTED_SCRAP_COLUMNS,
    INDONESIAN_MONTHS,
    parse_date_flexible
)

SHEET_DAILY = "Dashboard"
SHEET_FSB = "History FSB"
SHEET_SCRAP = "History SCRAP"

WORKSHEET_SCHEMAS = {
    SHEET_DAILY: {
        "table_name": "DASHBOARD",
        "default_header_row": 6,
        "default_start_col": 2,  # Col B
        "headers": [
            "WO", "Part Number", "Serial Number", "Costumer", "Status",
            "Progress", "PIR", "TAT", "Position", "Daily ", "Estimasi FSB", "Remark"
        ],
        "key_fields": ["WO"],
        "alias_to_norm": {
            "costumer": "Customer",
            "daily ": "Daily",
            "daily": "Daily",
            "catatan": "Remark"
        },
        "norm_to_actual": {
            "customer": "Costumer",
            "daily": "Daily "
        },
        "formula_cols": {
            "TAT": "=abs(H{row}-$E$5-INT((H{row}-$E$5)/7))"
        }
    },
    SHEET_FSB: {
        "table_name": "Table2",
        "default_header_row": 6,
        "default_start_col": 2,  # Col B
        "headers": [
            "WO", "Part Number", "Serial Number", "Costumer", "Status",
            "Progress", "PIR", "TAT", "Position", "Remark", "Date FSB"
        ],
        "key_fields": ["WO", "Serial Number", "Date FSB"],
        "alias_to_norm": {
            "costumer": "Customer",
            "catatan": "Remark"
        },
        "norm_to_actual": {
            "customer": "Costumer"
        },
        "formula_cols": {}
    },
    SHEET_SCRAP: {
        "table_name": "SCRAP",
        "default_header_row": 6,
        "default_start_col": 2,  # Col B
        "headers": [
            "WO", "Part Number", "Serial Number", "Costumer", "Status",
            "Progress", "PIR", "TAT", "Position", "Remark", "Date SCRAP", "Column 1"
        ],
        "key_fields": ["WO", "Serial Number", "Date SCRAP"],
        "alias_to_norm": {
            "costumer": "Customer",
            "catatan": "Remark"
        },
        "norm_to_actual": {
            "customer": "Costumer"
        },
        "formula_cols": {}
    }
}


def log_audit_event(action: str, wo: str = None, details: str = ""):
    """Logs local audit trail to audit_log.jsonl without network/cloud dependencies."""
    try:
        log_file = BASE_DIR / "audit_log.jsonl"
        entry = {
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "action": action,
            "wo": str(wo or ""),
            "details": details
        }
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(json.dumps(entry, ensure_ascii=False) + "\n")
    except Exception:
        pass


def parse_date_value(val):
    """Safely converts string/date/datetime into datetime.date or formatted representation."""
    d = parse_date_flexible(val)
    if d is not None:
        return d
    if val is not None and not pd.isna(val) and str(val).strip() not in ("", "-", "nan", "None", "NaT"):
        return str(val).strip()
    return None


def are_dates_equal(v1, v2) -> bool:
    """Compares two values as dates if possible."""
    if v1 is None or v2 is None:
        return False
    d1 = parse_date_value(v1)
    d2 = parse_date_value(v2)
    if isinstance(d1, datetime):
        d1 = d1.date()
    if isinstance(d2, datetime):
        d2 = d2.date()
    if isinstance(d1, date) and isinstance(d2, date):
        return d1 == d2
    return False


def normalize_clean_string(v) -> str:
    if v is None or pd.isna(v):
        return ""
    s = str(v).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    if s.lower() in ("nan", "none"):
        return ""
    return s


class ExcelRepository:
    def __init__(self, master_path: Path):
        self.master_path = Path(master_path)

    def set_master_path(self, new_path: Path):
        self.master_path = Path(new_path)

    def check_file_exists(self) -> bool:
        return self.master_path.exists() and self.master_path.stat().st_size > 0

    def resolve_sheet_name(self, sheet_name: str, available_sheets: list[str]) -> str:
        """
        Maps module names or legacy sheet names to the canonical master worksheet:
        - Daily Battery -> Dashboard
        - History FSB   -> History FSB
        - History SCRAP -> History SCRAP
        """
        clean_name = str(sheet_name or "").strip()
        if clean_name in ["Daily Battery", "Daily_Battery", "Dashboard"]:
            if "Dashboard" in available_sheets:
                return "Dashboard"
            if "Daily Battery" in available_sheets:
                return "Daily Battery"
            return SHEET_DAILY

        if clean_name in ["History FSB", "History_FSB", "FSB"]:
            for candidate in ["History FSB", "History_FSB", "FSB"]:
                if candidate in available_sheets:
                    return candidate
            return SHEET_FSB

        if clean_name in ["History SCRAP", "History_SCRAP", "SCRAP", "History SCRAP & FRA"]:
            for candidate in ["History SCRAP", "History_SCRAP", "SCRAP", "History SCRAP & FRA"]:
                if candidate in available_sheets:
                    return candidate
            return SHEET_SCRAP

        return clean_name

    def _get_table_info(self, ws, sheet_name: str):
        """
        Detects Excel Table or table range in the worksheet.
        Returns: (tbl_obj, s_row, s_col, e_row, e_col, actual_headers)
        """
        tbl = None
        schema = WORKSHEET_SCHEMAS.get(sheet_name, {})
        preferred_name = schema.get("table_name", "").upper()

        if ws.tables:
            for t in ws.tables.values():
                if preferred_name and t.name.upper() == preferred_name:
                    tbl = t
                    break
            if not tbl and len(ws.tables) > 0:
                tbl = list(ws.tables.values())[0]

        if tbl:
            s_row, s_col = coordinate_to_tuple(tbl.ref.split(':')[0])
            e_row, e_col = coordinate_to_tuple(tbl.ref.split(':')[1])
            headers = [str(ws.cell(s_row, c).value or "").strip() for c in range(s_col, e_col + 1)]

            # Check for any new rows added manually in Excel outside/below tbl.ref
            max_r = ws.max_row
            for r in range(e_row + 1, max_r + 1):
                has_val = any(ws.cell(r, col_i).value is not None and str(ws.cell(r, col_i).value).strip() != "" for col_i in range(s_col, e_col + 1))
                if has_val:
                    e_row = r
                else:
                    # Allow 1 empty row gap
                    if r + 1 <= max_r and any(ws.cell(r + 1, col_i).value is not None and str(ws.cell(r + 1, col_i).value).strip() != "" for col_i in range(s_col, e_col + 1)):
                        continue
                    break

            return tbl, s_row, s_col, e_row, e_col, headers

        # Fallback: scan for header row
        s_row = schema.get("default_header_row", 6)
        for r in range(1, min(16, ws.max_row + 1)):
            row_vals = [str(ws.cell(r, c).value or "").strip().lower() for c in range(1, min(20, ws.max_column + 1))]
            if any(target in row_vals for target in ["wo", "part number", "date fsb", "date scrap", "customer", "costumer"]):
                s_row = r
                break

        s_col = schema.get("default_start_col", 2)
        for c in range(1, 10):
            val_low = str(ws.cell(s_row, c).value or "").strip().lower()
            if val_low in ["wo", "part number", "date fsb", "date scrap"]:
                s_col = c
                break

        headers = []
        c = s_col
        while c <= ws.max_column:
            h_val = ws.cell(s_row, c).value
            if h_val is not None and str(h_val).strip():
                headers.append(str(h_val).strip())
                c += 1
            else:
                # peek next 2 cells
                if c + 1 <= ws.max_column and ws.cell(s_row, c + 1).value is not None:
                    headers.append("")
                    c += 1
                else:
                    break
        e_col = s_col + max(0, len(headers) - 1)

        # Detect last data row
        e_row = s_row
        for r in range(s_row + 1, ws.max_row + 1):
            has_val = any(ws.cell(r, col_i).value is not None and str(ws.cell(r, col_i).value).strip() != "" for col_i in range(s_col, e_col + 1))
            if has_val:
                e_row = r

        return None, s_row, s_col, e_row, e_col, headers

    def load_all_data(self) -> tuple[dict[str, pd.DataFrame], list[str]]:
        """
        Reads all sheets from master Excel without modifying the file.
        Uses evaluated values (data_only=True) so formulas like TAT are correctly populated.
        """
        if not self.check_file_exists():
            raise FileNotFoundError(f"Workbook master '{self.master_path}' tidak ditemukan di disk.")

        is_valid, errors, warnings = validate_workbook_schema(self.master_path)
        if not is_valid:
            raise ValueError(f"Validasi workbook master gagal:\n" + "\n".join(errors))

        wb = openpyxl.load_workbook(self.master_path, data_only=True, read_only=False)
        sheet_names = wb.sheetnames

        daily_sheet = self.resolve_sheet_name("Daily Battery", sheet_names)
        fsb_sheet = self.resolve_sheet_name("History FSB", sheet_names)
        scrap_sheet = self.resolve_sheet_name("History SCRAP", sheet_names)

        def read_sheet_dataframe(ws, sheet_key: str) -> pd.DataFrame:
            tbl, s_row, s_col, e_row, e_col, headers = self._get_table_info(ws, sheet_key)
            schema = WORKSHEET_SCHEMAS.get(sheet_key, {})
            alias_map = schema.get("alias_to_norm", {})

            # Map raw headers to normalized UI names
            norm_headers = []
            for h in headers:
                h_clean = h.strip()
                h_low = h.lower()
                if h_low in alias_map:
                    norm_headers.append(alias_map[h_low])
                elif h_clean == "Daily":
                    norm_headers.append("Daily")
                elif h_clean == "Catatan":
                    norm_headers.append("Remark")
                else:
                    norm_headers.append(h_clean)

            rows = []
            for r in range(s_row + 1, e_row + 1):
                vals = [ws.cell(r, c).value for c in range(s_col, e_col + 1)]
                # Check if row has at least one meaningful value
                if not any(v is not None and str(v).strip() != "" for v in vals):
                    continue

                row_dict = {}
                for idx, val in enumerate(vals):
                    if idx < len(norm_headers):
                        h_name = norm_headers[idx]
                        if not h_name:
                            continue
                        clean_v = normalize_clean_string(val)
                        row_dict[h_name] = clean_v

                # Valid row must have WO
                wo_val = row_dict.get("WO", "")
                if wo_val and wo_val.lower() != "wo":
                    rows.append(row_dict)

            df = pd.DataFrame(rows)
            # Ensure standard expected columns are present
            if sheet_key == SHEET_DAILY:
                for col in EXPECTED_DAILY_COLUMNS:
                    if col not in df.columns:
                        df[col] = ""
            elif sheet_key == SHEET_FSB:
                for col in EXPECTED_FSB_COLUMNS:
                    if col not in df.columns:
                        df[col] = ""
            elif sheet_key == SHEET_SCRAP:
                for col in EXPECTED_SCRAP_COLUMNS:
                    if col not in df.columns:
                        df[col] = ""
            return df

        df_daily = read_sheet_dataframe(wb[daily_sheet], SHEET_DAILY) if daily_sheet in wb.sheetnames else pd.DataFrame(columns=EXPECTED_DAILY_COLUMNS)
        df_fsb = read_sheet_dataframe(wb[fsb_sheet], SHEET_FSB) if fsb_sheet in wb.sheetnames else pd.DataFrame(columns=EXPECTED_FSB_COLUMNS)
        df_scrap = read_sheet_dataframe(wb[scrap_sheet], SHEET_SCRAP) if scrap_sheet in wb.sheetnames else pd.DataFrame(columns=EXPECTED_SCRAP_COLUMNS)

        wb.close()

        result = {
            "daily": df_daily,
            "fsb": df_fsb,
            "scrap": df_scrap,
            "sheet_names": sheet_names,
            "daily_sheet_name": daily_sheet,
            "fsb_sheet_name": fsb_sheet,
            "scrap_sheet_name": scrap_sheet
        }
        return result, warnings

    def _match_row(self, ws, s_row: int, s_col: int, e_col: int, headers: list[str], r_idx: int, identifier) -> bool:
        """Tests whether row `r_idx` matches the given identifier."""
        norm_row = {}
        for c_idx in range(s_col, e_col + 1):
            h_idx = c_idx - s_col
            if h_idx < len(headers):
                h_name = headers[h_idx].strip()
                raw_val = ws.cell(r_idx, c_idx).value
                cell_v = normalize_clean_string(raw_val)
                norm_row[h_name.lower()] = cell_v
                norm_row[f"_raw_{h_name.lower()}"] = raw_val
                if h_name.lower() == "costumer":
                    norm_row["customer"] = cell_v
                    norm_row["_raw_customer"] = raw_val
                if h_name.lower() in ("daily ", "daily"):
                    norm_row["daily"] = cell_v
                    norm_row["_raw_daily"] = raw_val

        if isinstance(identifier, dict):
            for k, expected_v in identifier.items():
                expected_clean = normalize_clean_string(expected_v).lower()
                if not expected_clean:
                    continue
                k_clean = k.strip().lower()
                actual_v = norm_row.get(k_clean, "").lower()
                if actual_v == expected_clean:
                    continue
                # If direct string match didn't match, check date equivalence
                raw_val = norm_row.get(f"_raw_{k_clean}")
                if are_dates_equal(expected_v, raw_val) or are_dates_equal(expected_v, actual_v):
                    continue
                return False
            return True
        else:
            expected_wo = normalize_clean_string(identifier).lower()
            return norm_row.get("wo", "").lower() == expected_wo

    def _apply_row_values(self, ws, row_idx: int, s_col: int, e_col: int, headers: list[str], data_dict: dict, sheet_name: str, prev_row_idx: int = None, is_update: bool = False):
        """Populates a row with data, formatting, styles, and formulas."""
        schema = WORKSHEET_SCHEMAS.get(sheet_name, {})
        formula_cols = schema.get("formula_cols", {})

        for col_idx in range(s_col, e_col + 1):
            h_idx = col_idx - s_col
            h_name = headers[h_idx].strip() if h_idx < len(headers) else ""
            cell = ws.cell(row=row_idx, column=col_idx)

            # In add mode, copy styling from previous data row or header
            if not is_update and prev_row_idx and prev_row_idx >= 6:
                prev_cell = ws.cell(row=prev_row_idx, column=col_idx)
                try:
                    cell.font = copy(prev_cell.font)
                    cell.border = copy(prev_cell.border)
                    cell.fill = copy(prev_cell.fill)
                    cell.alignment = copy(prev_cell.alignment)
                    cell.number_format = prev_cell.number_format
                except Exception:
                    pass

            # Resolve value from data_dict
            key_found = False
            val = None
            for candidate_key in [h_name, h_name.strip(), h_name.lower(), h_name.replace(" ", "")]:
                if candidate_key in data_dict:
                    val = data_dict[candidate_key]
                    key_found = True
                    break
            if not key_found:
                if h_name.lower() == "costumer" and ("Customer" in data_dict or "customer" in data_dict):
                    val = data_dict.get("Customer", data_dict.get("customer"))
                    key_found = True
                elif h_name.strip() == "Daily" and ("Daily" in data_dict or "daily" in data_dict):
                    val = data_dict.get("Daily", data_dict.get("daily"))
                    key_found = True
                elif h_name.lower() == "remark" and ("Catatan" in data_dict or "catatan" in data_dict or "Remark" in data_dict):
                    val = data_dict.get("Catatan", data_dict.get("catatan", data_dict.get("Remark")))
                    key_found = True

            # In update mode, if this column was NOT provided in data_dict, preserve existing cell value!
            if is_update and not key_found:
                continue

            # Formula check
            if h_name in formula_cols:
                # In Dashboard, TAT is dynamic Excel formula
                formula_template = formula_cols[h_name]
                cell.value = formula_template.format(row=row_idx)
                continue

            # Date formatting check
            if h_name in ["PIR", "Date FSB", "Date SCRAP"]:
                parsed_d = parse_date_value(val)
                if parsed_d is not None:
                    cell.value = parsed_d
                    cell.number_format = 'd " "mmmm" "yyyy'
                else:
                    cell.value = None if val is None or str(val).strip() == "" else str(val).strip()
                continue

            # TAT integer check (for History FSB / SCRAP)
            if h_name == "TAT":
                if val is not None and str(val).strip() != "":
                    try:
                        cell.value = int(float(str(val).strip()))
                    except (ValueError, TypeError):
                        cell.value = str(val).strip()
                else:
                    cell.value = 0
                continue

            # Numeric columns
            if h_name in ["Daily", "Daily ", "Estimasi FSB"]:
                if val is not None and str(val).strip().isdigit():
                    cell.value = int(str(val).strip())
                elif val is not None and str(val).strip() != "":
                    cell.value = str(val).strip()
                else:
                    cell.value = None
                continue

            # Standard string / empty
            if val is None or pd.isna(val) or str(val).strip() in ("", "-", "nan", "None"):
                cell.value = None
            else:
                cell.value = str(val).strip()

    def add_record(self, target_sheet_name: str, record: dict) -> tuple[bool, str]:
        """
        Appends a new record to the target worksheet.
        Expands the Excel Table boundary cleanly and preserves other sheets.
        Does NOT create a new sheet.
        """
        if not self.check_file_exists():
            return False, f"File master '{self.master_path}' tidak ditemukan."

        if is_file_locked(self.master_path):
            return False, f"File master '{self.master_path.name}' sedang dikunci oleh Microsoft Excel."

        temp_path = self.master_path.parent / f".tmp_{self.master_path.name}"
        backup_path = create_backup(self.master_path, reason=f"add_{target_sheet_name}")

        try:
            shutil.copy2(str(self.master_path), str(temp_path))
            wb = openpyxl.load_workbook(temp_path, data_only=False)

            actual_sheet_name = self.resolve_sheet_name(target_sheet_name, wb.sheetnames)
            if actual_sheet_name not in wb.sheetnames:
                wb.close()
                if temp_path.exists():
                    temp_path.unlink()
                return False, f"Sheet '{target_sheet_name}' tidak ditemukan di workbook master. Operasi dibatalkan tanpa membuat sheet baru."

            ws = wb[actual_sheet_name]
            tbl, s_row, s_col, e_row, e_col, headers = self._get_table_info(ws, actual_sheet_name)

            # Prevent duplicate insertion
            wo_to_add = normalize_clean_string(record.get("WO", ""))
            if wo_to_add:
                schema = WORKSHEET_SCHEMAS.get(actual_sheet_name, {})
                key_fields = schema.get("key_fields", ["WO"])
                for r in range(s_row + 1, e_row + 1):
                    if key_fields == ["WO"]:
                        if self._match_row(ws, s_row, s_col, e_col, headers, r, wo_to_add):
                            wb.close()
                            if temp_path.exists():
                                temp_path.unlink()
                            return False, f"Duplikasi terdeteksi: Nomor WO '{wo_to_add}' sudah ada di worksheet '{actual_sheet_name}'."
                    else:
                        composite = {k: record.get(k, "") for k in key_fields}
                        if self._match_row(ws, s_row, s_col, e_col, headers, r, composite):
                            wb.close()
                            if temp_path.exists():
                                temp_path.unlink()
                            return False, f"Duplikasi terdeteksi: Data dengan WO '{wo_to_add}' sudah terdaftar di worksheet '{actual_sheet_name}'."

            # Append row after last data row
            new_row_idx = e_row + 1
            ref_prev_row = e_row if e_row > s_row else s_row
            self._apply_row_values(ws, new_row_idx, s_col, e_col, headers, record, actual_sheet_name, prev_row_idx=ref_prev_row)

            # Expand table reference if sheet contains Excel Table
            if tbl:
                col_start_letter = get_column_letter(s_col)
                col_end_letter = get_column_letter(e_col)
                tbl.ref = f"{col_start_letter}{s_row}:{col_end_letter}{new_row_idx}"
                if tbl.autoFilter:
                    tbl.autoFilter.ref = tbl.ref

            wb.save(temp_path)
            wb.close()

            # Validate temp file
            test_wb = openpyxl.load_workbook(temp_path, read_only=True)
            if actual_sheet_name not in test_wb.sheetnames:
                test_wb.close()
                if temp_path.exists():
                    temp_path.unlink()
                return False, f"Verifikasi integritas gagal: Sheet '{actual_sheet_name}' tidak ditemukan di berkas sementara."
            test_wb.close()

            # Atomic replace
            os.replace(str(temp_path), str(self.master_path))
            return True, f"Data berhasil ditambahkan ke worksheet '{actual_sheet_name}'."

        except PermissionError:
            if temp_path.exists():
                try:
                    temp_path.unlink()
                except Exception:
                    pass
            return False, f"Akses ditolak: File '{self.master_path.name}' sedang dibuka dan dikunci oleh program lain (Excel)."
        except Exception as e:
            if temp_path.exists():
                try:
                    temp_path.unlink()
                except Exception:
                    pass
            return False, f"Gagal menambahkan data ke Excel: {e}"

    def update_record(self, target_sheet_name: str, identifier, updated_data: dict) -> tuple[bool, str]:
        """
        Updates an existing record in the target worksheet matching identifier.
        Preserves other rows, formatting, and formulas.
        """
        if not self.check_file_exists():
            return False, f"File master '{self.master_path}' tidak ditemukan."

        if is_file_locked(self.master_path):
            return False, f"File master '{self.master_path.name}' sedang dikunci oleh Microsoft Excel."

        temp_path = self.master_path.parent / f".tmp_{self.master_path.name}"
        create_backup(self.master_path, reason=f"update_{target_sheet_name}")

        try:
            shutil.copy2(str(self.master_path), str(temp_path))
            wb = openpyxl.load_workbook(temp_path, data_only=False)

            actual_sheet_name = self.resolve_sheet_name(target_sheet_name, wb.sheetnames)
            if actual_sheet_name not in wb.sheetnames:
                wb.close()
                if temp_path.exists():
                    temp_path.unlink()
                return False, f"Sheet '{target_sheet_name}' tidak ditemukan."

            ws = wb[actual_sheet_name]
            tbl, s_row, s_col, e_row, e_col, headers = self._get_table_info(ws, actual_sheet_name)

            target_row = None
            for r in range(s_row + 1, e_row + 1):
                if self._match_row(ws, s_row, s_col, e_col, headers, r, identifier):
                    target_row = r
                    break

            if not target_row:
                wb.close()
                if temp_path.exists():
                    temp_path.unlink()
                return False, f"Catatan yang ingin diperbarui tidak ditemukan di worksheet '{actual_sheet_name}'."

            # Update row values
            self._apply_row_values(ws, target_row, s_col, e_col, headers, updated_data, actual_sheet_name, prev_row_idx=target_row, is_update=True)

            wb.save(temp_path)
            wb.close()

            # Atomic replace
            os.replace(str(temp_path), str(self.master_path))
            return True, f"Data berhasil diperbarui pada worksheet '{actual_sheet_name}'."

        except PermissionError:
            if temp_path.exists():
                try:
                    temp_path.unlink()
                except Exception:
                    pass
            return False, f"Akses ditolak: File '{self.master_path.name}' sedang dibuka oleh Excel."
        except Exception as e:
            if temp_path.exists():
                try:
                    temp_path.unlink()
                except Exception:
                    pass
            return False, f"Gagal memperbarui data di Excel: {e}"

    def delete_record(self, target_sheet_name: str, identifier) -> tuple[bool, str]:
        """
        Deletes a specific record matching identifier from the target worksheet.
        Shrinks the table boundary cleanly and leaves other sheets untouched.
        """
        if not self.check_file_exists():
            return False, f"File master '{self.master_path}' tidak ditemukan."

        if is_file_locked(self.master_path):
            return False, f"File master '{self.master_path.name}' sedang dikunci oleh Microsoft Excel."

        temp_path = self.master_path.parent / f".tmp_{self.master_path.name}"
        create_backup(self.master_path, reason=f"delete_{target_sheet_name}")

        try:
            shutil.copy2(str(self.master_path), str(temp_path))
            wb = openpyxl.load_workbook(temp_path, data_only=False)

            actual_sheet_name = self.resolve_sheet_name(target_sheet_name, wb.sheetnames)
            if actual_sheet_name not in wb.sheetnames:
                wb.close()
                if temp_path.exists():
                    temp_path.unlink()
                return False, f"Sheet '{target_sheet_name}' tidak ditemukan."

            ws = wb[actual_sheet_name]
            tbl, s_row, s_col, e_row, e_col, headers = self._get_table_info(ws, actual_sheet_name)

            target_row = None
            max_search_row = max(e_row, ws.max_row)
            for r in range(s_row + 1, max_search_row + 1):
                if self._match_row(ws, s_row, s_col, e_col, headers, r, identifier):
                    target_row = r
                    break

            if not target_row:
                wb.close()
                if temp_path.exists():
                    temp_path.unlink()
                return False, f"Catatan yang ingin dihapus tidak ditemukan di worksheet '{actual_sheet_name}'."

            ws.delete_rows(target_row, 1)

            # Adjust table boundary
            new_e_row = max(s_row, e_row - 1)
            if tbl:
                col_start_letter = get_column_letter(s_col)
                col_end_letter = get_column_letter(e_col)
                tbl.ref = f"{col_start_letter}{s_row}:{col_end_letter}{new_e_row}"
                if tbl.autoFilter:
                    tbl.autoFilter.ref = tbl.ref

            # Clear trailing row cells if table shrank
            if new_e_row < e_row:
                for c in range(s_col, e_col + 1):
                    ws.cell(row=e_row, column=c).value = None

            # Re-evaluate dynamic formulas for any shifted rows (e.g. TAT in Dashboard)
            schema = WORKSHEET_SCHEMAS.get(actual_sheet_name, {})
            formula_cols = schema.get("formula_cols", {})
            if formula_cols:
                for r in range(target_row, new_e_row + 1):
                    for h_name, f_tpl in formula_cols.items():
                        if h_name in headers:
                            col_idx = s_col + headers.index(h_name)
                            ws.cell(row=r, column=col_idx).value = f_tpl.format(row=r)

            wb.save(temp_path)
            wb.close()

            # Atomic replace
            os.replace(str(temp_path), str(self.master_path))
            return True, f"Data berhasil dihapus dari worksheet '{actual_sheet_name}'."

        except PermissionError:
            if temp_path.exists():
                try:
                    temp_path.unlink()
                except Exception:
                    pass
            return False, f"Akses ditolak: File '{self.master_path.name}' sedang dibuka oleh Excel."
        except Exception as e:
            if temp_path.exists():
                try:
                    temp_path.unlink()
                except Exception:
                    pass
            return False, f"Gagal menghapus data dari Excel: {e}"

    def _route_unit_atomic(
        self,
        wo: str,
        target_sheet_alias: str,
        record_builder,
        audit_action: str,
        success_label: str,
        backup_reason: str
    ) -> tuple[bool, str]:
        clean_wo = normalize_clean_string(wo)
        if not clean_wo:
            return False, "Nomor Work Order (WO) tidak valid atau kosong."

        if not self.check_file_exists():
            return False, f"File master '{self.master_path}' tidak ditemukan."

        if is_file_locked(self.master_path):
            return False, f"File master '{self.master_path.name}' sedang dibuka dan dikunci oleh Microsoft Excel."

        temp_path = self.master_path.parent / f".tmp_{self.master_path.name}"
        create_backup(self.master_path, reason=backup_reason)

        try:
            shutil.copy2(str(self.master_path), str(temp_path))
            wb = openpyxl.load_workbook(temp_path, data_only=False)

            daily_sheet_name = self.resolve_sheet_name("Daily Battery", wb.sheetnames)
            target_sheet_name = self.resolve_sheet_name(target_sheet_alias, wb.sheetnames)

            if daily_sheet_name not in wb.sheetnames or target_sheet_name not in wb.sheetnames:
                wb.close()
                if temp_path.exists():
                    temp_path.unlink()
                return False, f"Worksheet yang dibutuhkan ('{daily_sheet_name}' atau '{target_sheet_name}') tidak ditemukan di workbook master."

            ws_daily = wb[daily_sheet_name]
            ws_target = wb[target_sheet_name]

            tbl_d, s_row_d, s_col_d, e_row_d, e_col_d, headers_d = self._get_table_info(ws_daily, daily_sheet_name)
            tbl_t, s_row_t, s_col_t, e_row_t, e_col_t, headers_t = self._get_table_info(ws_target, target_sheet_name)

            # 1. Locate source row in Dashboard
            source_row_idx = None
            max_search = max(e_row_d, ws_daily.max_row)
            for r in range(s_row_d + 1, max_search + 1):
                if self._match_row(ws_daily, s_row_d, s_col_d, e_col_d, headers_d, r, clean_wo):
                    source_row_idx = r
                    break

            if not source_row_idx:
                wb.close()
                if temp_path.exists():
                    temp_path.unlink()
                return False, f"Unit WO '{clean_wo}' tidak ditemukan pada worksheet Dashboard (Daily Battery aktif)."

            # Extract source row values into dictionary
            source_data = {}
            for c in range(s_col_d, e_col_d + 1):
                h_idx = c - s_col_d
                if h_idx < len(headers_d):
                    h_name = headers_d[h_idx].strip()
                    val = ws_daily.cell(row=source_row_idx, column=c).value
                    source_data[h_name] = val
                    if h_name.lower() == "costumer":
                        source_data["Customer"] = val
                    if h_name.lower() in ("daily ", "daily"):
                        source_data["Daily"] = val

            # 2. Check duplicate in target sheet
            for r in range(s_row_t + 1, max(e_row_t, ws_target.max_row) + 1):
                if self._match_row(ws_target, s_row_t, s_col_t, e_col_t, headers_t, r, clean_wo):
                    wb.close()
                    if temp_path.exists():
                        temp_path.unlink()
                    return False, f"Duplikasi terdeteksi: Unit WO '{clean_wo}' sudah tercatat di worksheet {target_sheet_alias}."

            # 3. Build target record via record_builder callback
            target_record, audit_details = record_builder(source_data, clean_wo)

            # 4. Append row to target sheet
            new_target_row = e_row_t + 1
            ref_prev = e_row_t if e_row_t > s_row_t else s_row_t
            self._apply_row_values(ws_target, new_target_row, s_col_t, e_col_t, headers_t, target_record, target_sheet_name, prev_row_idx=ref_prev)

            if tbl_t:
                col_start_t = get_column_letter(s_col_t)
                col_end_t = get_column_letter(e_col_t)
                tbl_t.ref = f"{col_start_t}{s_row_t}:{col_end_t}{new_target_row}"
                if tbl_t.autoFilter:
                    tbl_t.autoFilter.ref = tbl_t.ref

            # 5. Delete row from Dashboard
            ws_daily.delete_rows(source_row_idx, 1)
            new_e_row_d = max(s_row_d, e_row_d - 1)
            if tbl_d:
                col_start_d = get_column_letter(s_col_d)
                col_end_d = get_column_letter(e_col_d)
                tbl_d.ref = f"{col_start_d}{s_row_d}:{col_end_d}{new_e_row_d}"
                if tbl_d.autoFilter:
                    tbl_d.autoFilter.ref = tbl_d.ref

            # Clear trailing cell values in old bottom row
            if new_e_row_d < e_row_d:
                for c in range(s_col_d, e_col_d + 1):
                    ws_daily.cell(row=e_row_d, column=c).value = None

            # Re-evaluate dynamic formulas in Dashboard
            schema_d = WORKSHEET_SCHEMAS.get(daily_sheet_name, {})
            formula_cols = schema_d.get("formula_cols", {})
            if formula_cols:
                for r in range(source_row_idx, new_e_row_d + 1):
                    for h_name, f_tpl in formula_cols.items():
                        if h_name in headers_d:
                            col_i = s_col_d + headers_d.index(h_name)
                            ws_daily.cell(row=r, column=col_i).value = f_tpl.format(row=r)

            wb.save(temp_path)
            wb.close()

            # 6. Verify integrity of temp workbook before atomic replace
            test_wb = openpyxl.load_workbook(temp_path, read_only=True)
            if daily_sheet_name not in test_wb.sheetnames or target_sheet_name not in test_wb.sheetnames:
                test_wb.close()
                if temp_path.exists():
                    temp_path.unlink()
                return False, "Verifikasi integritas perpindahan data gagal."
            test_wb.close()

            # 7. Atomic replace
            os.replace(str(temp_path), str(self.master_path))
            log_audit_event(audit_action, wo=clean_wo, details=audit_details)
            return True, f"Unit WO '{clean_wo}' berhasil dipindahkan ke modul {success_label}!"

        except PermissionError:
            if temp_path.exists():
                try:
                    temp_path.unlink()
                except Exception:
                    pass
            return False, f"Akses ditolak: File '{self.master_path.name}' sedang dibuka di Microsoft Excel."
        except Exception as e:
            if temp_path.exists():
                try:
                    temp_path.unlink()
                except Exception:
                    pass
            return False, f"Gagal memproses otomatisasi {success_label}: {e}"

    def route_unit_to_fsb(self, wo: str, completion_date=None, extra_data: dict = None, completion_date_str=None, remark_str: str = "") -> tuple[bool, str]:
        """Moves a unit from 'Dashboard' (Daily Battery) to 'History FSB'."""
        if completion_date is None and completion_date_str is not None:
            completion_date = completion_date_str
        extra = dict(extra_data or {})
        if remark_str:
            extra["Remark"] = remark_str

        def build_fsb_record(source_data, clean_wo):
            fsb_date = parse_date_value(completion_date) or date.today()
            pir_val = source_data.get("PIR")
            pir_date = parse_date_value(pir_val)
            calculated_tat = source_data.get("TAT")
            if isinstance(pir_date, (datetime, date)) and isinstance(fsb_date, (datetime, date)):
                d_pir = pir_date.date() if isinstance(pir_date, datetime) else pir_date
                d_fsb = fsb_date.date() if isinstance(fsb_date, datetime) else fsb_date
                calculated_tat = max(1, (d_fsb - d_pir).days)

            record = {
                "WO": clean_wo,
                "Part Number": source_data.get("Part Number", ""),
                "Serial Number": source_data.get("Serial Number", ""),
                "Costumer": source_data.get("Costumer", source_data.get("Customer", "")),
                "Customer": source_data.get("Customer", source_data.get("Costumer", "")),
                "Status": source_data.get("Status", "Overhaul"),
                "Progress": "Final Charge",
                "PIR": pir_val,
                "TAT": calculated_tat,
                "Position": "UR",
                "Remark": source_data.get("Remark", ""),
                "Date FSB": fsb_date
            }
            record.update(extra)
            return record, f"Dipindahkan dari Daily Battery ke History FSB (Date FSB: {fsb_date})"

        return self._route_unit_atomic(
            wo=wo,
            target_sheet_alias="History FSB",
            record_builder=build_fsb_record,
            audit_action="AUTO_ROUTE_FINAL_CHARGE",
            success_label="History FSB",
            backup_reason=f"route_fsb_WO_{normalize_clean_string(wo)}"
        )

    def route_unit_to_scrap(self, wo: str, scrap_date=None, scrap_reason: str = None, extra_data: dict = None, scrap_date_str=None, reason_str: str = "", position_str: str = "") -> tuple[bool, str]:
        """Moves a unit from 'Dashboard' (Daily Battery) to 'History SCRAP'."""
        if scrap_date is None and scrap_date_str is not None:
            scrap_date = scrap_date_str
        if scrap_reason is None and reason_str:
            scrap_reason = reason_str
        extra = dict(extra_data or {})
        if position_str:
            extra["Position"] = position_str

        def build_scrap_record(source_data, clean_wo):
            sc_date = parse_date_value(scrap_date) or date.today()
            reason_text = str(scrap_reason or extra.get("Remark", source_data.get("Remark", "-"))).strip()
            pos_text = str(extra.get("Position", "SCRAP")).strip().upper()
            if pos_text not in ("SCRAP", "FRA"):
                pos_text = "SCRAP"

            record = {
                "WO": clean_wo,
                "Part Number": source_data.get("Part Number", ""),
                "Serial Number": source_data.get("Serial Number", ""),
                "Costumer": source_data.get("Costumer", source_data.get("Customer", "")),
                "Customer": source_data.get("Customer", source_data.get("Costumer", "")),
                "Status": source_data.get("Status", "Overhaul"),
                "Progress": "SCRAP",
                "PIR": source_data.get("PIR"),
                "TAT": source_data.get("TAT", 0),
                "Position": pos_text,
                "Remark": reason_text,
                "Date SCRAP": sc_date,
                "Column 1": ""
            }
            record.update(extra)
            record["Position"] = pos_text
            record["Remark"] = reason_text
            return record, f"Dipindahkan dari Daily Battery ke History SCRAP (Alasan: {reason_text})"

        return self._route_unit_atomic(
            wo=wo,
            target_sheet_alias="History SCRAP",
            record_builder=build_scrap_record,
            audit_action="AUTO_ROUTE_SCRAP",
            success_label="History SCRAP",
            backup_reason=f"route_scrap_WO_{normalize_clean_string(wo)}"
        )

    def save_sheet_atomic(self, df: pd.DataFrame, target_sheet_name: str, columns: list[str]) -> tuple[bool, str]:
        """
        Safely synchronizes a dataframe into the target sheet inside master workbook.
        Strictly preserves rows 1..5 (titles, today formula), existing table headers, and styles.
        Never creates new sheets.
        """
        if not self.check_file_exists():
            return False, f"File master '{self.master_path}' tidak ditemukan."

        if is_file_locked(self.master_path):
            return False, f"File '{self.master_path.name}' sedang dibuka dan dikunci oleh Microsoft Excel."

        temp_path = self.master_path.parent / f".tmp_{self.master_path.name}"
        create_backup(self.master_path, reason=f"sync_{target_sheet_name}")

        try:
            shutil.copy2(str(self.master_path), str(temp_path))
            wb = openpyxl.load_workbook(temp_path, data_only=False)

            actual_sheet_name = self.resolve_sheet_name(target_sheet_name, wb.sheetnames)
            if actual_sheet_name not in wb.sheetnames:
                wb.close()
                if temp_path.exists():
                    temp_path.unlink()
                return False, f"Worksheet target '{target_sheet_name}' tidak ditemukan di workbook master."

            ws = wb[actual_sheet_name]
            tbl, s_row, s_col, e_row, e_col, headers = self._get_table_info(ws, actual_sheet_name)

            # Delete existing data rows only (rows after header row)
            if e_row > s_row:
                ws.delete_rows(s_row + 1, e_row - s_row)

            # Write rows from dataframe
            curr_row = s_row
            for _, r in df.iterrows():
                curr_row += 1
                row_dict = r.to_dict()
                self._apply_row_values(ws, curr_row, s_col, e_col, headers, row_dict, actual_sheet_name, prev_row_idx=s_row)

            # Update table reference
            if tbl:
                col_start_letter = get_column_letter(s_col)
                col_end_letter = get_column_letter(e_col)
                new_e_row = max(s_row, curr_row)
                tbl.ref = f"{col_start_letter}{s_row}:{col_end_letter}{new_e_row}"

            wb.save(temp_path)
            wb.close()

            # Verify integrity
            test_wb = openpyxl.load_workbook(temp_path, read_only=True)
            if actual_sheet_name not in test_wb.sheetnames:
                test_wb.close()
                if temp_path.exists():
                    temp_path.unlink()
                return False, f"Verifikasi gagal: Sheet '{actual_sheet_name}' hilang di file sementara."
            test_wb.close()

            # Atomic replacement
            os.replace(str(temp_path), str(self.master_path))
            return True, f"Perubahan berhasil disimpan ke worksheet '{actual_sheet_name}'."

        except PermissionError:
            if temp_path.exists():
                try:
                    temp_path.unlink()
                except Exception:
                    pass
            return False, f"File '{self.master_path.name}' sedang dibuka dan dikunci oleh Microsoft Excel."
        except Exception as e:
            if temp_path.exists():
                try:
                    temp_path.unlink()
                except Exception:
                    pass
            return False, f"Terjadi kesalahan saat menyimpan data: {e}"

    def export_copy_safe(self, destination_path: Path, sheet_to_export: str = None) -> tuple[bool, str]:
        """
        Exports a standalone copy of the workbook.
        Strictly prevents overwriting the active master workbook.
        """
        dest = Path(destination_path).resolve()
        if dest == self.master_path.resolve():
            return False, "File ekspor tidak boleh menimpa file Excel master aktif!"

        if not self.check_file_exists():
            return False, "Workbook master tidak ditemukan untuk diekspor."

        try:
            shutil.copy2(str(self.master_path), str(dest))
            return True, f"Salinan laporan berhasil diekspor ke: {dest.name}"
        except Exception as e:
            return False, f"Gagal mengekspor salinan: {e}"
