"""
Configuration Manager for Aircraft Battery Maintenance Monitoring System
Handles application-level settings, master Excel path, and backup policies.
"""

import sys
import json
from pathlib import Path

if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent.resolve()
else:
    BASE_DIR = Path(__file__).parent.resolve()

def get_resource_path(relative_path: str) -> Path:
    """
    Returns absolute path to resource.
    Supports PyInstaller one-file/one-dir bundle (sys._MEIPASS) and development environment.
    """
    if hasattr(sys, "_MEIPASS"):
        bundle_path = Path(sys._MEIPASS) / relative_path
        if bundle_path.exists():
            return bundle_path
    return BASE_DIR / relative_path

CONFIG_FILE = BASE_DIR / "config.json"

DEFAULT_CONFIG = {
    "master_excel_filename": "Daily Battery.xlsx",
    "backup_folder": "backups",
    "backup_retention_count": 30,
    "auto_backup_on_save": True,
    "detect_external_conflicts": True,
    "conflict_check_tolerance_seconds": 1.0
}

def load_app_config() -> dict:
    """Loads config from config.json or initializes with defaults."""
    cfg = DEFAULT_CONFIG.copy()
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                if isinstance(saved, dict):
                    cfg.update(saved)
        except Exception as e:
            print(f"[Config] Warning: Failed to parse {CONFIG_FILE.name}: {e}. Using defaults.")
    else:
        save_app_config(cfg)
    return cfg

def save_app_config(cfg: dict) -> bool:
    """Saves config dictionary to config.json safely."""
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=4, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"[Config] Error saving config: {e}")
        return False

def get_master_excel_path() -> Path:
    """Returns the absolute Path object pointing to the official Master Excel file."""
    cfg = load_app_config()
    filename_or_path = cfg.get("master_excel_filename", "Daily Battery.xlsx")
    p = Path(filename_or_path)
    if not p.is_absolute():
        p = BASE_DIR / p
    # If the configured file does not exist (e.g. app moved to another laptop or path altered),
    # smartly try finding the filename inside BASE_DIR, or fallback to default Daily Battery.xlsx
    if not p.exists():
        candidate_local = BASE_DIR / p.name
        if candidate_local.exists():
            return candidate_local
        candidate_default = BASE_DIR / "Daily Battery.xlsx"
        if candidate_default.exists():
            return candidate_default
    return p

def set_master_excel_path(new_path: Path) -> bool:
    """Updates master Excel file path in config."""
    cfg = load_app_config()
    p = Path(new_path).resolve()
    # If it is inside BASE_DIR, store as relative or name for portability
    try:
        rel = p.relative_to(BASE_DIR)
        cfg["master_excel_filename"] = str(rel)
    except ValueError:
        cfg["master_excel_filename"] = str(p)
    return save_app_config(cfg)
