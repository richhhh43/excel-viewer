@echo off
set REPO=C:\Users\rich_\OneDrive\Documents\GitHub\excel-viewer

set PATH=C:\Program Files\Git\cmd;%PATH%

cd /d "%REPO%"
python "%REPO%\publish_sheet.py"
pause
