
@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1
set PYTHONUNBUFFERED=1
where python >nul 2>&1
if %errorlevel%==0 (
    python "%~dp0run_bulk_copy_ui.py"
    goto :eof
)
where py >nul 2>&1
if %errorlevel%==0 (
    py "%~dp0run_bulk_copy_ui.py"
    goto :eof
)
echo No se encontró Python en PATH. Instala Python o ejecuta manualmente:
echo   %~dp0run_bulk_copy_ui.py
pause
