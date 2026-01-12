@echo off
REM Build script for all AfterShell exfil tools
REM Compiles both Debug and Release versions of all projects

setlocal

REM Get the directory where this script is located
set "SCRIPT_DIR=%~dp0"

echo ============================================
echo  AfterShell Exfil Tools - Build All
echo ============================================
echo.

REM ============================================
REM Build ScreenCapture
REM ============================================
call "%SCRIPT_DIR%ScreenCapture\build-screencapture.bat"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] ScreenCapture build failed!
    exit /b 1
)

REM ============================================
REM Build OutlookExporter
REM ============================================
echo.
call "%SCRIPT_DIR%OutlookExporter\build-outlookexporter.bat"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] OutlookExporter build failed!
    exit /b 1
)

echo.
echo ============================================
echo  All Projects Built Successfully!
echo ============================================
echo.

endlocal
