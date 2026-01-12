@echo off
REM Quick build and run script for all AfterShell exfil tools
REM Builds both projects in Release mode

setlocal

REM Get the directory where this script is located
set "SCRIPT_DIR=%~dp0"

echo ============================================
echo  AfterShell Exfil Tools - Build and Run All
echo ============================================
echo.

REM ============================================
REM Build and Run ScreenCapture
REM ============================================
start "ScreenCapture" cmd /c "call "%SCRIPT_DIR%ScreenCapture\build-and-run-screencapture.bat""

REM ============================================
REM Build and Run OutlookExporter
REM ============================================
echo.
start "OutlookExporter" cmd /c "call "%SCRIPT_DIR%OutlookExporter\build-and-run-outlookexporter.bat""

echo.
echo ============================================
echo  Both applications are building and running!
echo ============================================
echo.

endlocal
