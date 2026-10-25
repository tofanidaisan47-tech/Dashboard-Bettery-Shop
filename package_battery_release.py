import os
import shutil
import json
import zipfile

src_dir = r"c:\Users\ASUS\.gemini\antigravity-ide\scratch\battery_shop_tcc4"
dist_src = os.path.join(src_dir, "dist", "Dashboard_Battery_TCC2")
final_dir = os.path.join(src_dir, "Dashboard_Battery_Final")

print("1. Updating Dashboard_Battery_Final folder...")

# Replace _internal and exe
internal_dir = os.path.join(final_dir, "_internal")
if os.path.exists(internal_dir):
    shutil.rmtree(internal_dir)
shutil.copytree(os.path.join(dist_src, "_internal"), internal_dir)

shutil.copy2(
    os.path.join(dist_src, "Dashboard_Battery_TCC2.exe"),
    os.path.join(final_dir, "Dashboard_Battery_TCC2.exe"),
)

# Assets
assets_dir = os.path.join(final_dir, "assets")
if os.path.exists(assets_dir):
    shutil.rmtree(assets_dir)
shutil.copytree(os.path.join(src_dir, "assets"), assets_dir)

# Dropdown config
shutil.copy2(
    os.path.join(src_dir, "dropdown_config.json"),
    os.path.join(final_dir, "dropdown_config.json"),
)

# Daily Battery.xlsx
shutil.copy2(
    os.path.join(src_dir, "Daily Battery.xlsx"),
    os.path.join(final_dir, "Daily Battery.xlsx"),
)

# Config.json (portable configuration pointing to local Daily Battery.xlsx)
cfg = {
    "master_excel_filename": "Daily Battery.xlsx",
    "backup_folder": "backups",
    "backup_retention_count": 30,
    "auto_backup_on_save": True,
    "detect_external_conflicts": True,
    "conflict_check_tolerance_seconds": 1.0,
}
with open(os.path.join(final_dir, "config.json"), "w", encoding="utf-8") as f:
    json.dump(cfg, f, indent=4)

# Batch launcher
bat_content = "@echo off\r\ncd /d \"%~dp0\"\r\nstart \"\" \"Dashboard_Battery_TCC2.exe\"\r\n"
with open(os.path.join(final_dir, "Buka_Dashboard_Battery.bat"), "w", encoding="ascii") as f:
    f.write(bat_content)

# Desktop shortcut creator script
shortcut_bat = """@echo off
title Buat Shortcut Desktop - Battery Shop TCC-2
color 0A
set TARGET_DIR=%~dp0
set TARGET_EXE=%TARGET_DIR%Dashboard_Battery_TCC2.exe
set ICON_FILE=%TARGET_DIR%assets\\app_icon.ico

echo ======================================================================
echo    MEMBUAT SHORTCUT DESKTOP: DASHBOARD BATTERY SHOP TCC-2
echo ======================================================================
powershell -NoProfile -Command "$ws = New-Object -ComObject WScript.Shell; $d = [Environment]::GetFolderPath('Desktop'); $s = $ws.CreateShortcut([System.IO.Path]::Combine($d, 'Dashboard Battery TCC-2.lnk')); $s.TargetPath = '%TARGET_EXE%'; $s.WorkingDirectory = '%TARGET_DIR%'; if (Test-Path '%ICON_FILE%') { $s.IconLocation = '%ICON_FILE%' }; $s.Save()"

echo.
echo  [SUKSES] Shortcut 'Dashboard Battery TCC-2' berhasil dibuat di Desktop!
echo  Anda sekarang dapat membuka aplikasi langsung dari layar desktop laptop.
echo.
pause
"""
with open(os.path.join(final_dir, "Buat_Shortcut_Desktop.bat"), "w", encoding="ascii") as f:
    f.write(shortcut_bat)

# Panduan penggunaan
guide = """================================================================================
 BATTERY SHOP TCC-2 | PT GMF AEROASIA Tbk
 Aircraft Battery Maintenance Dashboard & Operational Monitoring System
================================================================================

PANDUAN PENGGUNAAN DI LAPTOP PERUSAHAAN (100% OFFLINE - STANDALONE)

1. CARA MENJALANKAN APLIKASI:
   - Klik ganda (double-click) pada file 'Dashboard_Battery_TCC2.exe' atau 'Buka_Dashboard_Battery.bat'.
   - (Opsional) Klik ganda 'Buat_Shortcut_Desktop.bat' untuk membuat ikon pintasan langsung di Desktop Windows laptop Anda.
   - Aplikasi akan langsung terbuka secara mandiri (100% Standalone).
   - Laptop kantor TIDAK MEMERLUKAN instalasi Python atau software tambahan lainnya.

2. SUMBER DATA MASTER EXCEL:
   - File default adalah 'Daily Battery.xlsx' yang berada di folder ini.
   - Anda dapat membuka & mengedit file ini langsung di Microsoft Excel kapan saja.
   - Jika Anda memiliki file Excel yang diupdate dari HP / laptop lain:
     a. Buka aplikasi, klik menu '📁 Excel ▾' di header atas.
     b. Pilih '📁 Pilih / Import Master Workbook Lain...'.
     c. Pilih file Excel Anda, lalu klik 'Ya' pada konfirmasi validasi.
     d. Dashboard akan otomatis memuat dan beralih ke file tersebut sebagai master aktif.

3. KESELAMATAN & CADANGAN OTOMATIS (BACKUPS):
   - Setiap kali terjadi penambahan/perubahan/penghapusan data di dashboard,
     sistem otomatis membuat cadangan aman di folder 'backups/'.
   - Jika terjadi kendala pada file Excel, Anda dapat memulihkan cadangan
     melalui menu 'Pengaturan' > 'Pulihkan Data dari Cadangan'.

4. PINTASAN KEYBOARD (SHORTCUTS):
   - [Double-Click Sel] : Edit data langsung atau pilih dropdown di tabel
   - [Delete]           : Hapus baris yang sedang dipilih
   - [F5]               : Refresh data dari file Excel aktif
   - [F1]               : Buka Glosarium & Panduan Alur Status Bengkel

================================================================================
PT GMF AeroAsia Tbk - Safety • Quality • Reliability
================================================================================
"""
with open(os.path.join(final_dir, "PANDUAN_PENGGUNAAN.txt"), "w", encoding="utf-8") as f:
    f.write(guide)

# Backups folder
backups_dir = os.path.join(final_dir, "backups")
if os.path.exists(backups_dir):
    shutil.rmtree(backups_dir)
os.makedirs(backups_dir, exist_ok=True)

print("2. Deleting old zip file...")
old_zip = os.path.join(src_dir, "Dashboard_Battery_Final.zip")
if os.path.exists(old_zip):
    os.remove(old_zip)
    print("Deleted old zip:", old_zip)

print("3. Creating new zip archive...")
with zipfile.ZipFile(old_zip, "w", zipfile.ZIP_DEFLATED) as zf:
    for root, dirs, files in os.walk(final_dir):
        for file in files:
            full_path = os.path.join(root, file)
            rel_path = os.path.relpath(full_path, os.path.dirname(final_dir))
            zf.write(full_path, rel_path)

print("New zip created successfully:", old_zip)
print("Size (MB):", round(os.path.getsize(old_zip) / (1024 * 1024), 2))

# Also copy to D:\Semester 5\Daisan Tofani\GMF\PROJECT\Dashboard Battery\kirim ini
gmf_dest = r"D:\Semester 5\Daisan Tofani\GMF\PROJECT\Dashboard Battery\kirim ini"
if os.path.exists(gmf_dest):
    dest_zip = os.path.join(gmf_dest, "Dashboard_Battery_Final.zip")
    shutil.copy2(old_zip, dest_zip)
    print("Copied to folder kirim ini:", dest_zip)

# Also copy to D:\Semester 5\Daisan Tofani\GMF\PROJECT\Dashboard Battery
gmf_root = r"D:\Semester 5\Daisan Tofani\GMF\PROJECT\Dashboard Battery"
if os.path.exists(gmf_root):
    dest_zip2 = os.path.join(gmf_root, "Dashboard_Battery_Final.zip")
    shutil.copy2(old_zip, dest_zip2)
    print("Copied to folder Dashboard Battery:", dest_zip2)
