@echo off
REM Quick build and run script for Outlook Data Exporter

setlocal

REM Get the directory where this script is located
set "SCRIPT_DIR=%~dp0"

echo Building Outlook Data Exporter...
dotnet build "%SCRIPT_DIR%OutlookExporter.csproj"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo Build successful!
    echo.
    echo Running the application...
    echo.
    dotnet run --project "%SCRIPT_DIR%OutlookExporter.csproj" -- %*
) else (
    echo.
    echo Build failed!
    exit /b 1
)

endlocal
