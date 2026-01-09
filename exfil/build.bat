@echo off
REM Build script for Outlook Data Exporter
REM Compiles both Debug and Release versions

setlocal

REM Get the directory where this script is located
set "SCRIPT_DIR=%~dp0"

echo ============================================
echo  Outlook Data Exporter - Build Script
echo ============================================
echo.

REM Clean previous builds
echo Cleaning previous builds...
dotnet clean "%SCRIPT_DIR%OutlookExporter.csproj" -v q >nul 2>&1

REM Build Debug version
echo.
echo Building Debug version...
dotnet build "%SCRIPT_DIR%OutlookExporter.csproj" -c Debug -v q

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Debug build failed!
    exit /b 1
)
echo [OK] Debug build successful

REM Build Release version
echo.
echo Building Release version...
dotnet build "%SCRIPT_DIR%OutlookExporter.csproj" -c Release -v q

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Release build failed!
    exit /b 1
)
echo [OK] Release build successful

REM Show output locations
echo.
echo ============================================
echo  Build Complete!
echo ============================================
echo.
echo Output locations:
echo   Debug:   %SCRIPT_DIR%bin\Debug\net8.0-windows\OutlookExporter.exe
echo   Release: %SCRIPT_DIR%bin\Release\net8.0-windows\OutlookExporter.exe
echo.

endlocal
