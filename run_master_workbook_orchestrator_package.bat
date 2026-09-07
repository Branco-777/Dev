@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
pushd "%SCRIPT_DIR%"
if errorlevel 1 (
    echo ERROR: Could not access the script directory:
    echo "%SCRIPT_DIR%"
    pause
    exit /b 2
)
set "PYTHON=%LOCALAPPDATA%\Just\master-workbook-orchestrator\Scripts\python.exe"
if not exist "%PYTHON%" set "PYTHON=python"
if "%~1"=="" set "MASTER=%SCRIPT_DIR%Master Spreadsheet Template.xlsx"
if not "%~1"=="" set "MASTER=%~1"
if not exist "%MASTER%" (
    echo ERROR: Master workbook not found:
    echo "%MASTER%"
    popd
    pause
    exit /b 2
)
echo Running consolidated master workbook orchestrator...
"%PYTHON%" "%SCRIPT_DIR%master_workbook_orchestrator_merged.py" --master "%MASTER%" --overwrite
set "EXIT_CODE=%ERRORLEVEL%"
echo.
if %EXIT_CODE% equ 0 (
    echo Orchestrator completed. Check the configured output folders.
) else (
    echo ERROR: Orchestrator failed with exit code %EXIT_CODE%.
)
popd
pause
exit /b %EXIT_CODE%
