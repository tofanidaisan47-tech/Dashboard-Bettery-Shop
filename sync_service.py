"""
Sync Service for Aircraft Battery Maintenance Monitoring System
Handles fingerprinting, external modification conflict detection, and lock testing.
"""

import hashlib
import os
from datetime import datetime
from pathlib import Path

def get_file_fingerprint(file_path: Path) -> dict:
    """
    Computes fingerprint of a file based on size, mtime, and SHA-256 hash.
    Used to detect if external modifications occurred in Microsoft Excel.
    """
    p = Path(file_path)
    if not p.exists():
        return {
            "exists": False,
            "path": str(p),
            "size": 0,
            "mtime": 0.0,
            "mtime_str": "-",
            "sha256": ""
        }

    stat = p.stat()
    mtime = stat.st_mtime
    size = stat.st_size

    # Fast hash calculation
    hasher = hashlib.sha256()
    try:
        with open(p, "rb") as f:
            # Read first and last 64KB for speed, or entire file if small (<1MB)
            if size <= 1024 * 1024:
                hasher.update(f.read())
            else:
                hasher.update(f.read(65536))
                f.seek(max(0, size - 65536))
                hasher.update(f.read(65536))
        digest = hasher.hexdigest()
    except Exception:
        digest = f"size_{size}_mtime_{mtime}"

    dt_str = datetime.fromtimestamp(mtime).strftime("%d %b %Y, %H:%M:%S")
    return {
        "exists": True,
        "path": str(p.resolve()),
        "size": size,
        "mtime": mtime,
        "mtime_str": dt_str,
        "sha256": digest
    }

def check_external_conflict(arg1, arg2, tolerance_sec: float = 1.0) -> tuple[bool, str]:
    """
    Detects if the file was modified externally after loaded_fingerprint was captured.
    Accepts either (file_path, loaded_fingerprint) or (loaded_fingerprint, file_path).
    Returns: (has_conflict: bool, message: str)
    """
    if isinstance(arg1, dict):
        loaded_fingerprint = arg1
        current_file_path = Path(arg2)
    else:
        current_file_path = Path(arg1)
        loaded_fingerprint = arg2

    if not loaded_fingerprint or not isinstance(loaded_fingerprint, dict) or not loaded_fingerprint.get("exists"):
        return False, "Tidak ada data riwayat sinkronisasi sebelumnya."

    curr = get_file_fingerprint(current_file_path)
    if not curr["exists"]:
        return True, "File master tidak ditemukan di disk!"

    # Check if hash changed or mtime has shifted beyond tolerance
    time_diff = abs(curr["mtime"] - loaded_fingerprint.get("mtime", 0.0))
    hash_changed = curr["sha256"] != loaded_fingerprint.get("sha256", "")

    if hash_changed or (time_diff > tolerance_sec):
        msg = (
            f"Perhatian: File Excel telah dimodifikasi oleh program lain (misal: Microsoft Excel)\n"
            f"pada {curr['mtime_str']} (terakhir dibaca dashboard: {loaded_fingerprint.get('mtime_str')}).\n\n"
            "Menyimpan data sekarang berisiko menimpa perubahan eksternal tersebut!\n"
            "Disarankan untuk menekan tombol 'Sync / Refresh Data' terlebih dahulu."
        )
        return True, msg

    return False, "File sinkron dan belum berubah."

def has_external_conflict(arg1, arg2, tolerance_sec: float = 1.0) -> bool:
    """Returns True if file was modified externally after fingerprint was recorded."""
    is_conflict, _ = check_external_conflict(arg1, arg2, tolerance_sec)
    return is_conflict

def check_file_locked(file_path: Path) -> tuple[bool, str]:
    """
    Checks if the Excel file is currently locked/in-use by Microsoft Excel or WPS Office.
    Returns: (is_locked: bool, message: str)
    """
    p = Path(file_path)
    if not p.exists():
        return False, "File tidak ada."

    # 1. Check for MS Excel lock file (~$filename.xlsx)
    lock_file = p.parent / f"~${p.name}"
    if lock_file.exists():
        return True, f"File lock '~${p.name}' terdeteksi (sedang dibuka di Microsoft Excel)."

    # 2. Try opening file in append mode to test OS lock
    try:
        with open(p, "a+b"):
            pass
        return False, "File dapat diakses secara normal."
    except PermissionError:
        return True, f"Akses ditolak: File '{p.name}' sedang dibuka dan dikunci oleh program lain (Excel/WPS)."
    except Exception as e:
        return True, f"Tidak dapat mengakses file: {e}"

def is_file_locked(file_path: Path) -> bool:
    """Returns True if the file is currently locked by another program."""
    locked, _ = check_file_locked(file_path)
    return locked
