@echo off
title Buat Shortcut Desktop - Battery Shop TCC-2
color 0A
set TARGET_DIR=%~dp0
set TARGET_EXE=%TARGET_DIR%Dashboard_Battery_TCC2.exe
set ICON_FILE=%TARGET_DIR%assets\app_icon.ico

echo ======================================================================
echo    MEMBUAT SHORTCUT DESKTOP: DASHBOARD BATTERY SHOP TCC-2
echo ======================================================================
powershell -NoProfile -Command "$ws = New-Object -ComObject WScript.Shell; $d = [Environment]::GetFolderPath('Desktop'); $s = $ws.CreateShortcut([System.IO.Path]::Combine($d, 'Dashboard Battery TCC-2.lnk')); $s.TargetPath = '%TARGET_EXE%'; $s.WorkingDirectory = '%TARGET_DIR%'; if (Test-Path '%ICON_FILE%') { $s.IconLocation = '%ICON_FILE%' }; $s.Save()"

echo.
echo  [SUKSES] Shortcut 'Dashboard Battery TCC-2' berhasil dibuat di Desktop!
echo  Anda sekarang dapat membuka aplikasi langsung dari layar desktop laptop.
echo.
pause
