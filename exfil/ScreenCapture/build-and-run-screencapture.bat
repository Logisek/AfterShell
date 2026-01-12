@echo off
REM Quick build and run script for Screen Capture Utility

setlocal

REM Get the directory where this script is located
set "SCRIPT_DIR=%~dp0"

echo Publishing Screen Capture Utility (Debug configuration)...
dotnet publish "%SCRIPT_DIR%ScreenCapture.csproj" -c Debug -r win-x64 --self-contained true -v q

if %ERRORLEVEL% EQU 0 (
    echo.
    echo Publish successful!
    echo Single-file executable: %SCRIPT_DIR%bin\Debug\net8.0-windows\win-x64\publish\ScreenCapture.exe
    echo.
    echo Running the application...
    echo.
    "%SCRIPT_DIR%bin\Debug\net8.0-windows\win-x64\publish\ScreenCapture.exe" %*
) else (
    echo.
    echo Publish failed!
    exit /b 1
)

endlocal
