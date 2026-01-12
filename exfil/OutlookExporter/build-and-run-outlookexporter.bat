@echo off
REM Quick build and run script for Outlook Data Exporter

setlocal

REM Get the directory where this script is located
set "SCRIPT_DIR=%~dp0"

echo Publishing Outlook Data Exporter (Debug configuration)...
dotnet publish "%SCRIPT_DIR%OutlookExporter.csproj" -c Debug -r win-x64 --self-contained true -v q

if %ERRORLEVEL% EQU 0 (
    echo.
    echo Publish successful!
    echo Single-file executable: %SCRIPT_DIR%bin\Debug\net8.0-windows\win-x64\publish\OutlookExporter.exe
    echo.
    echo Running the application...
    echo.
    "%SCRIPT_DIR%bin\Debug\net8.0-windows\win-x64\publish\OutlookExporter.exe" %*
) else (
    echo.
    echo Publish failed!
    exit /b 1
)

endlocal
