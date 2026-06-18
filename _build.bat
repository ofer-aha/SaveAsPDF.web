@echo off
cd /d "%~dp0"
echo Building SaveAsPDF webpack bundle...
call npm run build
if %errorlevel% neq 0 (
    echo.
    echo BUILD FAILED - check errors above
    pause
    exit /b 1
)
echo.
echo Build complete. Output in C:\APPS\SaveAsPDF\wwwroot
pause
