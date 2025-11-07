@echo off
REM Build script untuk Windows

echo ==========================================
echo Building DOCX Placeholder Replacer
echo ==========================================

REM Activate virtual environment if exists
if exist "venv\Scripts\activate.bat" (
    call venv\Scripts\activate.bat
)

REM Clean previous builds
echo Cleaning previous builds...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist

REM Build using PyInstaller
echo Building executable...
pyinstaller docx-replacer.spec

REM Check if build succeeded
if %ERRORLEVEL% EQU 0 (
    echo.
    echo ==========================================
    echo Build completed successfully!
    echo ==========================================
    echo.
    echo Executable: dist\DOCX-Replacer.exe
    echo.
    echo To run:
    echo   dist\DOCX-Replacer.exe
    echo.
    echo You can distribute the 'dist' folder to users.
    echo They don't need Python installed!
) else (
    echo.
    echo ==========================================
    echo Build failed! Check errors above.
    echo ==========================================
    exit /b 1
)
