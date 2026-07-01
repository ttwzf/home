@echo off
setlocal
chcp 65001 >nul

set "SCRIPT_DIR=%~dp0"
set "BUNDLED_PY=C:\Users\photos\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

if exist "%BUNDLED_PY%" (
  "%BUNDLED_PY%" "%SCRIPT_DIR%optimize_process_table.py"
) else (
  py "%SCRIPT_DIR%optimize_process_table.py"
)

pause
