"""
Validation Module for Aircraft Battery Maintenance Monitoring System
Validates workbook schema, sheet structure, column headers, and data integrity constraints.
"""

import re
from datetime import datetime, date
from pathlib import Path
import pandas as pd
import openpyxl

MONTHS_ID = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"]
MONTHS_FULL_ID = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
]
DAYS_FULL_ID = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]

INDONESIAN_MONTHS = {
    "jan": 1, "januari": 1,
    "feb": 2, "februari": 2,
    "mar": 3, "maret": 3,
    "apr": 4, "april": 4,
    "mei": 5, "may": 5,
    "jun": 6, "juni": 6, "june": 6,
    "jul": 7, "juli": 7, "july": 7,
    "agu": 8, "agustus": 8, "aug": 8, "august": 8,
    "sep": 9, "september": 9,
    "okt": 10, "oktober": 10, "oct": 10, "october": 10,
    "nov": 11, "november": 11,
    "des": 12, "desember": 12, "dec": 12, "december": 12,
}

def parse_date_flexible(val):
    """Centralized flexible date parser supporting datetime, Timestamp, standard strings, and Indonesian dates."""
    if val is None or pd.isna(val):
        return None
    if isinstance(val, (datetime, pd.Timestamp)):
        return val.date()
    if isinstance(val, date):
        return val
    val_str = str(val).strip()
    if not val_str or val_str.lower() in ("", "nan", "none", "nat", "-", "null", "00:00:00"):
        return None
    
    parts = val_str.split()
    if len(parts) == 3:
        d, m, y = parts
        m_lower = m.lower()
        if m_lower in INDONESIAN_MONTHS and d.isdigit() and y.isdigit():
            try:
                return date(int(y), INDONESIAN_MONTHS[m_lower], int(d))
            except ValueError:
                pass
    for fmt in (
        "%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y",
        "%d %b %Y", "%d %B %Y", "%d/%m/%Y %H:%M:%S", "%d-%m-%Y %H:%M:%S"
    ):
        try:
            return datetime.strptime(val_str, fmt).date()
        except Exception:
            pass
    try:
        dt = pd.to_datetime(val_str, errors='coerce', format='mixed')
        if pd.notna(dt):
            return dt.date()
    except Exception:
        pass
    return None

to_date_object = parse_date_flexible

def format_date_with_month_name(val) -> str:
    """Formats date/datetime/string into standard Indonesian format (e.g. '15 Oktober 2026')."""
    if val is None or pd.isna(val):
        return ""
    val_str = str(val).strip()
    if not val_str or val_str.lower() in ("", "nan", "none", "nat"):
        return ""
    if val_str == "-":
        return "-"
    if re.match(r"^\d{1,2}$", val_str):
        return val_str
    if isinstance(val, (datetime, date, pd.Timestamp)):
        idx = val.month - 1
        return f"{val.day} {MONTHS_FULL_ID[idx]} {val.year}"
    m_name = re.match(r"^(\d{1,2})[- ]+([A-Za-z]+)[- ]+(\d{4})$", val_str)
    if m_name:
        d, m_text, y = m_name.groups()
        return f"{int(d)} {m_text.capitalize()} {y}"
    try:
        dt = pd.to_datetime(val_str, errors='coerce', format='mixed')
        if pd.notna(dt):
            idx = dt.month - 1
            return f"{dt.day} {MONTHS_FULL_ID[idx]} {dt.year}"
    except Exception:
        pass
    return val_str

EXPECTED_DAILY_COLUMNS = [
    "WO", "Part Number", "Serial Number", "Customer", "Status",
    "Progress", "PIR", "TAT", "Position", "Daily", "Estimasi FSB", "Remark"
]

EXPECTED_FSB_COLUMNS = [
    "Date FSB", "WO", "Part Number", "Serial Number", "Customer",
    "Status", "Progress", "TAT", "Position", "Remark"
]

EXPECTED_SCRAP_COLUMNS = [
    "Date SCRAP", "WO", "Part Number", "Serial Number", "Customer",
    "Status", "TAT", "Position", "Remark"
]

def validate_workbook_schema(file_path: Path) -> tuple[bool, list[str], list[str]]:
    """
    Validates the structure of the master Excel workbook.
    Returns: (is_valid, errors, warnings)
    """
    p = Path(file_path)
    errors = []
    warnings = []

    if not p.exists():
        errors.append(f"File workbook master tidak ditemukan di lokasi: {p}")
        return False, errors, warnings

    if p.stat().st_size == 0:
        errors.append(f"File workbook master kosong (0 bytes): {p.name}")
        return False, errors, warnings

    try:
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        sheets = wb.sheetnames

        # Check Daily Battery sheet
        daily_sheet_found = False
        for s in ["Daily Battery", "Daily_Battery", "Dashboard"]:
            if s in sheets:
                daily_sheet_found = True
                break
        if not daily_sheet_found:
            errors.append(f"Sheet utama 'Daily Battery' tidak ditemukan. Sheet yang ada: {sheets}")

        # Check optional history sheets
        fsb_found = any(s in sheets for s in ["History FSB", "History_FSB", "FSB"])
        if not fsb_found:
            warnings.append("Sheet 'History FSB' tidak ditemukan di workbook.")

        scrap_found = any(s in sheets for s in ["History SCRAP", "History_SCRAP", "SCRAP", "History SCRAP & FRA"])
        if not scrap_found:
            warnings.append("Sheet 'History SCRAP' tidak ditemukan di workbook.")

        wb.close()
    except Exception as e:
        errors.append(f"Gagal membaca format workbook Excel: {e}")
        return False, errors, warnings

    is_valid = len(errors) == 0
    return is_valid, errors, warnings


def inspect_workbook_summary(file_path: Path) -> dict:
    """
    Performs full structure & content summary inspection for Import / Welcome preview.
    Returns dictionary with validation status, sheet counts, and schema details.
    """
    is_valid, errors, warnings = validate_workbook_schema(file_path)
    summary = {
        "is_valid": is_valid,
        "errors": errors,
        "warnings": warnings,
        "daily_count": 0,
        "fsb_count": 0,
        "scrap_count": 0,
        "sheets": [],
        "file_name": Path(file_path).name,
        "file_size_kb": round(Path(file_path).stat().st_size / 1024, 1) if Path(file_path).exists() else 0
    }
    if not is_valid:
        return summary

    try:
        from excel_repository import ExcelRepository
        repo = ExcelRepository(file_path)
        data, repo_warns = repo.load_all_data()
        summary["warnings"].extend(repo_warns)
        summary["daily_count"] = len(data.get("daily", []))
        summary["fsb_count"] = len(data.get("fsb", []))
        summary["scrap_count"] = len(data.get("scrap", []))
        summary["sheets"] = data.get("sheet_names", [])
    except Exception as e:
        summary["is_valid"] = False
        summary["errors"].append(f"Gagal memuat isi sheet: {e}")

    return summary

def validate_battery_unit_record(
    record: dict,
    existing_df: pd.DataFrame = None,
    mode: str = "add",
    original_wo: str = None,
    existing_wos: set | list = None,
    is_edit: bool = False
) -> tuple[bool, str]:
    """
    Validates a single battery record before insertion or update.
    Returns: (is_valid: bool, error_message: str)
    """
    if is_edit:
        mode = "edit"

    errors = []
    wo = str(record.get("WO", "")).strip()
    sn = str(record.get("Serial Number", "")).strip()
    pn = str(record.get("Part Number", "")).strip()
    cust = str(record.get("Customer", "")).strip()
    tat = record.get("TAT", 0)

    # Required fields
    if not wo:
        errors.append("Nomor WO (Work Order) wajib diisi.")
    if not sn:
        errors.append("Nomor Serial Number (S/N) wajib diisi.")
    if not pn:
        errors.append("Part Number (P/N) wajib diisi.")
    if not cust:
        errors.append("Customer (Maskapai) wajib diisi.")

    # TAT validation
    if tat is not None and str(tat).strip() != "":
        try:
            int_tat = int(float(str(tat).strip()))
            if int_tat < 0:
                errors.append("Nilai TAT (Turn Around Time) tidak boleh negatif.")
        except (ValueError, TypeError):
            errors.append("Nilai TAT harus berupa angka bulat.")

    # Collect existing WOs set (case-insensitive)
    wos_set = set()
    if existing_wos is not None:
        wos_set = {str(x).strip().lower() for x in existing_wos if pd.notna(x)}
    elif existing_df is not None and not existing_df.empty and "WO" in existing_df.columns:
        wos_set = {str(x).strip().lower() for x in existing_df["WO"].dropna()}

    # Duplicate WO check in active WIP (mode == 'add')
    if mode == "add":
        if wo.lower() in wos_set:
            errors.append(f"Nomor WO '{wo}' sudah terdaftar aktif di sheet Daily Battery! Harap gunakan nomor WO unik.")

    # In mode == 'edit', if WO is changed, ensure the new WO does not collide with another existing record
    elif mode == "edit":
        orig = str(original_wo or "").strip().lower()
        if wo.lower() != orig:
            if orig in wos_set:
                wos_set.remove(orig)
            if wo.lower() in wos_set:
                errors.append(f"Nomor WO '{wo}' sudah digunakan oleh unit lain di sheet Daily Battery!")

    is_valid = len(errors) == 0
    err_msg = "• " + "\n• ".join(errors) if errors else ""
    return is_valid, err_msg


def validate_fsb_record(
    record: dict,
    existing_df: pd.DataFrame = None,
    mode: str = "add",
    original_identifier = None
) -> tuple[bool, str]:
    """Validates an FSB record before addition or update."""
    errors = []
    wo = str(record.get("WO", "")).strip()
    sn = str(record.get("Serial Number", "")).strip()
    pn = str(record.get("Part Number", "")).strip()
    cust = str(record.get("Customer", "")).strip()
    date_fsb = str(record.get("Date FSB", "")).strip()
    tat = record.get("TAT", 0)

    if not wo:
        errors.append("Nomor WO (Work Order) wajib diisi.")
    if not sn:
        errors.append("Nomor Serial Number (S/N) wajib diisi.")
    if not pn:
        errors.append("Part Number (P/N) wajib diisi.")
    if not cust:
        errors.append("Customer (Maskapai) wajib diisi.")
    if not date_fsb:
        errors.append("Date FSB wajib diisi.")

    if tat is not None and str(tat).strip() != "":
        try:
            int_tat = int(float(str(tat).strip()))
            if int_tat < 0:
                errors.append("Nilai TAT tidak boleh negatif.")
        except (ValueError, TypeError):
            errors.append("Nilai TAT harus berupa angka bulat.")

    if mode == "add" and existing_df is not None and not existing_df.empty and "WO" in existing_df.columns:
        match_wo = existing_df["WO"].astype(str).str.strip().str.lower() == wo.lower()
        match_sn = existing_df["Serial Number"].astype(str).str.strip().str.lower() == sn.lower() if "Serial Number" in existing_df.columns else True
        match_dt = existing_df["Date FSB"].astype(str).str.strip().str.lower() == date_fsb.lower() if "Date FSB" in existing_df.columns else True
        if (match_wo & match_sn & match_dt).any():
            errors.append(f"Catatan FSB dengan WO '{wo}', S/N '{sn}', dan tanggal '{date_fsb}' sudah tercatat.")

    is_valid = len(errors) == 0
    err_msg = "• " + "\n• ".join(errors) if errors else ""
    return is_valid, err_msg


def validate_scrap_record(
    record: dict,
    existing_df: pd.DataFrame = None,
    mode: str = "add",
    original_identifier = None
) -> tuple[bool, str]:
    """Validates a SCRAP record before addition or update."""
    errors = []
    wo = str(record.get("WO", "")).strip()
    sn = str(record.get("Serial Number", "")).strip()
    pn = str(record.get("Part Number", "")).strip()
    cust = str(record.get("Customer", "")).strip()
    date_scrap = str(record.get("Date SCRAP", "")).strip()
    tat = record.get("TAT", 0)

    if not wo:
        errors.append("Nomor WO (Work Order) wajib diisi.")
    if not sn:
        errors.append("Nomor Serial Number (S/N) wajib diisi.")
    if not pn:
        errors.append("Part Number (P/N) wajib diisi.")
    if not cust:
        errors.append("Customer (Maskapai) wajib diisi.")
    if not date_scrap:
        errors.append("Date SCRAP wajib diisi.")

    if tat is not None and str(tat).strip() != "":
        try:
            int_tat = int(float(str(tat).strip()))
            if int_tat < 0:
                errors.append("Nilai TAT tidak boleh negatif.")
        except (ValueError, TypeError):
            errors.append("Nilai TAT harus berupa angka bulat.")

    if mode == "add" and existing_df is not None and not existing_df.empty and "WO" in existing_df.columns:
        match_wo = existing_df["WO"].astype(str).str.strip().str.lower() == wo.lower()
        match_sn = existing_df["Serial Number"].astype(str).str.strip().str.lower() == sn.lower() if "Serial Number" in existing_df.columns else True
        match_dt = existing_df["Date SCRAP"].astype(str).str.strip().str.lower() == date_scrap.lower() if "Date SCRAP" in existing_df.columns else True
        if (match_wo & match_sn & match_dt).any():
            errors.append(f"Catatan SCRAP dengan WO '{wo}', S/N '{sn}', dan tanggal '{date_scrap}' sudah tercatat.")

    is_valid = len(errors) == 0
    err_msg = "• " + "\n• ".join(errors) if errors else ""
    return is_valid, err_msg

