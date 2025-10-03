@echo off
rem --- CONFIGURA ESTO ---
set "TARGET_DIR=C:\Users\User\PycharmProjects\Contabilidad"
set "VENV_DIR=venv"
set "CMD_TO_LOAD=python manage.py runserver"   rem <-- comando que quieres tener listo

rem --- Copiar el comando al portapapeles (PowerShell) ---
powershell -NoProfile -Command "Set-Clipboard -Value '%CMD_TO_LOAD%'"

rem --- Abrir cmd, ir a la carpeta y activar el virtualenv ---
rem Usamos call para ejecutar el activate.bat y luego dejar la consola en el venv
start "" cmd.exe /k "cd /d \"%TARGET_DIR%\" && if exist \"%VENV_DIR%\\Scripts\\activate.bat\" ( call \"%VENV_DIR%\\Scripts\\activate.bat\" ) else ( echo AVISO: no se encontro %VENV_DIR%\\Scripts\\activate.bat )"

echo.
echo El virtualenv (si existe) fue activado en %TARGET_DIR%.
echo El comando fue copiado al portapapeles. Presiona Ctrl+V y Enter para ejecutarlo.
pause