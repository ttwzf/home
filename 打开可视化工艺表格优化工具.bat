@echo off
setlocal
chcp 65001 >nul

set "SCRIPT_DIR=%~dp0"
set "BUNDLED_PY=C:\Users\photos\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

if exist "%BUNDLED_PY%" (
  start "" "%BUNDLED_PY%" "%SCRIPT_DIR%visual_process_optimizer.py"
) else (
  start "" py "%SCRIPT_DIR%visual_process_optimizer.py"
)
