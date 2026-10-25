"""
Backup Manager for Aircraft Battery Maintenance Monitoring System
Handles automated, timestamped local workbook backups, retention rotation, and recovery.
"""

import os
import shutil
from datetime import datetime
from pathlib import Path
from config import BASE_DIR, load_app_config

def get_backup_directory() -> Path:
    """Returns the backup directory path, ensuring it exists."""
    cfg = load_app_config()
    folder_name = cfg.get("backup_folder", "backups")
    p = Path(folder_name)
    if not p.is_absolute():
        p = BASE_DIR / p
    p.mkdir(parents=True, exist_ok=True)
    return p

def create_backup(source_file: Path, reason: str = "pre_save") -> Path | None:
    """
    Creates a verified local backup of source_file before risky modifications.
    Returns the Path of the backup created, or None if source file does not exist.
    """
    source_path = Path(source_file)
    if not source_path.exists():
        return None

    backup_dir = get_backup_directory()
    now_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    stem = source_path.stem.replace(" ", "_")
    clean_reason = reason.replace(" ", "_")
    backup_filename = f"{stem}_backup_{now_str}_{clean_reason}.xlsx"
    backup_path = backup_dir / backup_filename

    # If same second filename collision happens, add a suffix
    counter = 1
    while backup_path.exists():
        backup_filename = f"{stem}_backup_{now_str}_{clean_reason}_{counter}.xlsx"
        backup_path = backup_dir / backup_filename
        counter += 1

    try:
        shutil.copy2(str(source_path), str(backup_path))
        # Verify backup file size is greater than 0
        if backup_path.exists() and backup_path.stat().st_size > 0:
            enforce_retention_policy()
            return backup_path
        else:
            if backup_path.exists():
                backup_path.unlink()
            return None
    except Exception as e:
        print(f"[Backup] Error creating backup for {source_path.name}: {e}")
        return None

def list_backups() -> list[dict]:
    """
    Returns list of existing backups sorted descending (newest first).
    Each item contains: filename, path, size_kb, created_at, created_at_str.
    """
    backup_dir = get_backup_directory()
    backups = []
    for item in backup_dir.glob("*.xlsx"):
        if item.is_file() and not item.name.startswith("~$") and not item.name.startswith(".tmp"):
            stat = item.stat()
            dt = datetime.fromtimestamp(stat.st_mtime)
            backups.append({
                "filename": item.name,
                "path": item,
                "size_kb": round(stat.st_size / 1024, 1),
                "created_at": dt,
                "created_at_str": dt.strftime("%d %b %Y, %H:%M:%S")
            })
    backups.sort(key=lambda x: (x["created_at"], x["filename"]), reverse=True)
    return backups

def enforce_retention_policy():
    """
    Deletes oldest backups exceeding the retention threshold.
    NEVER deletes the last remaining valid backup.
    """
    cfg = load_app_config()
    max_count = int(cfg.get("backup_retention_count", 30))
    if max_count <= 0:
        return

    all_backups = list_backups()
    if len(all_backups) > max_count:
        # Keep the newest max_count, delete the rest, but ensure at least 1 backup remains
        to_delete = all_backups[max_count:]
        for b in to_delete:
            try:
                b["path"].unlink()
                print(f"[Backup] Rotated old backup: {b['filename']}")
            except Exception as e:
                print(f"[Backup] Could not rotate {b['filename']}: {e}")

def restore_backup(backup_path: Path, target_path: Path) -> tuple[bool, str]:
    """
    Safely restores a backup file to target master location.
    Creates a pre-restore safety snapshot of the target first if it exists.
    """
    src = Path(backup_path)
    dst = Path(target_path)
    if not src.exists():
        return False, f"File backup '{src.name}' tidak ditemukan."

    try:
        # Safety: snapshot current target before replacing
        if dst.exists():
            create_backup(dst, reason="pre_restore_safety")

        shutil.copy2(str(src), str(dst))
        return True, f"Berhasil memulihkan data dari '{src.name}' ke '{dst.name}'."
    except Exception as e:
        return False, f"Gagal memulihkan backup: {e}"
