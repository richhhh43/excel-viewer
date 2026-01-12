@echo on
setlocal

REM Go to this repo folder
cd /d "%~dp0"

echo Running publisher from: %cd%
echo.

REM Run and SHOW output live (no redirect)
python publish_sheet.py

echo.
echo Finished running publish_sheet.py
pause
