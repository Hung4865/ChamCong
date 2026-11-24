@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0lay_cong.ps1" %*
if errorlevel 1 (
    echo Lay RAW that bai. Xem thong bao loi phia tren.
)
pause
