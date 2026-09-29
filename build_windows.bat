@echo off
REM ==== Build brands_crawler.exe locally (run once, on Windows) ====
REM Needs Python installed (python.org, tick "Add Python to PATH").
cd /d "%~dp0"
python -m pip install --upgrade pip
python -m pip install pyinstaller -r requirements.txt
pyinstaller --onefile --icon icon.ico --name brands_crawler crawler.py
echo.
echo ============================================================
echo Done. Your program is:  dist\brands_crawler.exe
echo Copy it into the folder with your .xlsx files,
echo then double-click it whenever you want to update.
echo ============================================================
pause
