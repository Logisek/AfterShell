@echo off
REM Build script for Screen Capture Utility
REM Compiles both Debug and Release versions

setlocal

REM Get the directory where this script is located
set "SCRIPT_DIR=%~dp0"

echo ============================================
echo  Screen Capture Utility - Build Script
echo ============================================
echo.

REM Clean previous builds
echo Cleaning previous builds...
dotnet clean "%SCRIPT_DIR%ScreenCapture.csproj" -v q >nul 2>&1

REM Publish Debug version (single-file executable)
echo.
echo Publishing Debug version (single-file)...
dotnet publish "%SCRIPT_DIR%ScreenCapture.csproj" -c Debug -r win-x64 --self-contained true -v q

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Debug publish failed!
    exit /b 1
)
echo [OK] Debug publish successful

REM Publish Release version (single-file executable)
echo.
echo Publishing Release version (single-file)...
dotnet publish "%SCRIPT_DIR%ScreenCapture.csproj" -c Release -r win-x64 --self-contained true -v q

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Release publish failed!
    exit /b 1
)
echo [OK] Release publish successful

REM Show output locations
echo.
echo ============================================
echo  Build Complete!
echo ============================================
echo.
echo Single-file executables (no DLLs required):
echo   Debug:   %SCRIPT_DIR%bin\Debug\net8.0-windows\win-x64\publish\ScreenCapture.exe
echo   Release: %SCRIPT_DIR%bin\Release\net8.0-windows\win-x64\publish\ScreenCapture.exe
echo.

endlocal
