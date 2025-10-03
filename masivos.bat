@echo off
setlocal

REM === CONFIGURA ESTO ===
set "WORKDIR=%~dp0"
REM set "WORKDIR=C:\Users\User\PycharmProjects\Contabilidad"

set "CMD_TO_RUN=python bulk_copy_sharepoint_graph.py --mode masiva --excel "masivo.xlsx" --same-file "masivo.pdf" --sheet "Hoja1""

REM Abrir una nueva ventana de CMD con todo listo
start "VENV - %WORKDIR%" cmd /k ^
 "cd /d ""%WORKDIR%"" ^
  && if exist .venv\Scripts\activate.bat (call .venv\Scripts\activate.bat) else (echo No se encontro .venv\Scripts\activate.bat) ^
  && doskey go=%CMD_TO_RUN% ^
  && echo. ^
  && echo Comando preparado: %CMD_TO_RUN% ^
  && echo Alias creado: escribe go y presiona Enter ^
  && echo (Tambien lo copie al portapapeles) ^
  && echo %CMD_TO_RUN%|clip"
