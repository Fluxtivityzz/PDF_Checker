@echo off
setlocal
cd /d "%~dp0"

set "PY_CMD=python"
where py >nul 2>nul
if not errorlevel 1 set "PY_CMD=py -3"

%PY_CMD% -c "import fitz, openpyxl" >nul 2>nul
if errorlevel 1 (
    echo Missing dependencies. Installing pymupdf and openpyxl...
    %PY_CMD% -m pip install pymupdf openpyxl
    if errorlevel 1 (
        echo.
        echo Dependency installation failed.
        pause
        exit /b 1
    )
)

%PY_CMD% "%~dp0check_pdf_gui.py"
if errorlevel 1 (
    echo.
    echo Application exited with an error.
    pause
    exit /b 1
)

endlocal
