@echo on
setlocal

echo -----------------------------
echo publish.bat starting...
echo Script folder: %~dp0
echo Current folder BEFORE cd: %cd%
echo -----------------------------

cd /d "%~dp0"
echo Current folder AFTER cd: %cd%

echo.
echo Checking python...
where python
python --version
echo.

echo Running publish_sheet.py...
python publish_sheet.py > publish_log.txt 2>&1

echo.
echo -------- publish_log.txt --------
type publish_log.txt
echo -------- end log --------
echo.

echo Done. Log saved at: %cd%\publish_log.txt
pause
