@echo off
setlocal
cd /d "%~dp0"

set "APP_NAME=PDF_Checker"
set "PY_CMD=python"
where py >nul 2>nul
if not errorlevel 1 set "PY_CMD=py -3"

echo Checking build dependencies...
%PY_CMD% -c "import fitz, openpyxl, PyInstaller" >nul 2>nul
if errorlevel 1 (
    echo Installing pymupdf, openpyxl and pyinstaller...
    %PY_CMD% -m pip install pymupdf openpyxl pyinstaller
    if errorlevel 1 (
        echo.
        echo Dependency installation failed.
        pause
        exit /b 1
    )
)

echo.
echo Building %APP_NAME%.exe...
%PY_CMD% -m PyInstaller ^
    --noconfirm ^
    --clean ^
    --windowed ^
    --name "%APP_NAME%" ^
    --add-data "check_rules.json;." ^
    "%~dp0check_pdf_gui.py"

if errorlevel 1 (
    echo.
    echo Build failed.
    pause
    exit /b 1
)

echo.
echo Build complete:
echo %~dp0dist\%APP_NAME%\%APP_NAME%.exe
pause
endlocal
