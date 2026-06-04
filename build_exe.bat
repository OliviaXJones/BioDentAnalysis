@echo off
echo ============================================================
echo  BioDent EXE Builder
echo ============================================================
echo.

:: Check for PyInstaller
python -c "import PyInstaller" >nul 2>&1
if %errorlevel% neq 0 (
    echo PyInstaller not found. Installing...
    pip install pyinstaller
    if %errorlevel% neq 0 (
        echo ERROR: Could not install PyInstaller. Aborting.
        pause
        exit /b 1
    )
)

:: Clean previous build artifacts
if exist build   rmdir /s /q build
if exist dist    rmdir /s /q dist
if exist BioDent.spec del /q BioDent.spec

echo.
echo Building BioDent.exe (this may take a minute)...
echo.

python -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name BioDent ^
    --hidden-import PyQt5.sip ^
    --hidden-import openpyxl ^
    --hidden-import pandas ^
    --hidden-import FKBP5_BioDent_Pipeline ^
    --hidden-import SingleStudy_BioDent_Pipeline ^
    --hidden-import BioDent_Utils ^
    --collect-all PyQt5 ^
    BioDent_Main.py

if %errorlevel% neq 0 (
    echo.
    echo ERROR: Build failed. See output above for details.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Build complete!
echo  EXE location: dist\BioDent.exe
echo.
echo  NOTE: studies_config.json must be in the same folder as
echo  BioDent.exe when distributing.
echo ============================================================
echo.
pause
