@echo off
REM Script to start Edge with remote debugging enabled for URL monitoring
REM This allows the test case capture tool to monitor browser URLs

echo Starting Microsoft Edge with remote debugging enabled...
echo.
echo The test case capture tool will be able to monitor URLs from this Edge instance.
echo.
echo Press Ctrl+C to close Edge when done.
echo.

REM Find Edge installation
set EDGE_PATH=
if exist "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" (
    set EDGE_PATH=C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe
) else if exist "C:\Program Files\Microsoft\Edge\Application\msedge.exe" (
    set EDGE_PATH=C:\Program Files\Microsoft\Edge\Application\msedge.exe
) else (
    echo Edge not found in default locations.
    echo Please edit this file and set EDGE_PATH to your Edge installation path.
    pause
    exit /b 1
)

REM Start Edge with remote debugging on port 9222
start "" "%EDGE_PATH%" --remote-debugging-port=9222 --user-data-dir="%TEMP%\edge_debug_profile"

echo.
echo Edge started! You can now use the test case capture tool.
echo Navigate to https://qa-exchange.doceree.com to start capturing test cases.
echo.
pause
