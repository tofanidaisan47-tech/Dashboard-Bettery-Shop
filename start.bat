@echo off
title Battery Shop TCC-2 (GMF AeroAsia) - Desktop Dashboard GUI Launcher
color 0B

echo ======================================================================
echo    GMF AEROASIA - BATTERY SHOP UNIT TCC-2 DESKTOP DASHBOARD GUI
echo ======================================================================
echo  [INFO] Menjalankan Aplikasi Desktop Dashboard PC/Laptop (100%% Offline)...
echo  [INFO] Framework: CustomTkinter & ttk.Treeview
echo  [INFO] Data Source: Master Excel Workbook (Dinamis / Sesuai Import)
echo  [INFO] Membuka antarmuka grafis... Harap tunggu sejenak.
echo ======================================================================

cd /d "%~dp0"
python app.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  [ERROR] Terjadi kesalahan saat menjalankan aplikasi.
    echo  pip install customtkinter pandas openpyxl tkcalendar matplotlib pillow
    echo.
    pause
)
