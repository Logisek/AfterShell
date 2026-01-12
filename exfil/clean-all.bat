@echo off
REM Clean script for all AfterShell exfil tools
REM Removes all build artifacts (bin and obj directories)

setlocal

REM Get the directory where this script is located
set "SCRIPT_DIR=%~dp0"

echo ============================================
echo  AfterShell Exfil Tools - Clean All
echo ============================================
echo.

REM ============================================
REM Clean ScreenCapture
REM ============================================
echo Cleaning ScreenCapture...
if exist "%SCRIPT_DIR%ScreenCapture\bin" (
    echo   Removing ScreenCapture\bin...
    rmdir /s /q "%SCRIPT_DIR%ScreenCapture\bin"
)
if exist "%SCRIPT_DIR%ScreenCapture\obj" (
    echo   Removing ScreenCapture\obj...
    rmdir /s /q "%SCRIPT_DIR%ScreenCapture\obj"
)
echo   ScreenCapture cleaned.

REM ============================================
REM Clean OutlookExporter
REM ============================================
echo.
echo Cleaning OutlookExporter...
if exist "%SCRIPT_DIR%OutlookExporter\bin" (
    echo   Removing OutlookExporter\bin...
    rmdir /s /q "%SCRIPT_DIR%OutlookExporter\bin"
)
if exist "%SCRIPT_DIR%OutlookExporter\obj" (
    echo   Removing OutlookExporter\obj...
    rmdir /s /q "%SCRIPT_DIR%OutlookExporter\obj"
)
echo   OutlookExporter cleaned.

echo.
echo ============================================
echo  All Projects Cleaned Successfully!
echo ============================================
echo.

endlocal
