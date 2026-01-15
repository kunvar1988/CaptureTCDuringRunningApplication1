@echo off
REM Script to start Chrome with remote debugging enabled for URL monitoring
REM This allows the test case capture tool to monitor browser URLs

echo Starting Chrome with remote debugging enabled...
echo.
echo The test case capture tool will be able to monitor URLs from this Chrome instance.
echo.
echo Press Ctrl+C to close Chrome when done.
echo.

REM Find Chrome installation
set CHROME_PATH=
if exist "C:\Program Files\Google\Chrome\Application\chrome.exe" (
    set CHROME_PATH=C:\Program Files\Google\Chrome\Application\chrome.exe
) else if exist "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" (
    set CHROME_PATH=C:\Program Files (x86)\Google\Chrome\Application\chrome.exe
) else (
    echo Chrome not found in default locations.
    echo Please edit this file and set CHROME_PATH to your Chrome installation path.
    pause
    exit /b 1
)

REM Start Chrome with remote debugging on port 9222
start "" "%CHROME_PATH%" --remote-debugging-port=9222 --user-data-dir="%TEMP%\chrome_debug_profile"

echo.
echo Chrome started! You can now use the test case capture tool.
echo Navigate to https://qa-exchange.doceree.com to start capturing test cases.
echo.
pause
