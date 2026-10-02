@echo off
rem Builds "dist\Statement to Excel.exe" - a single file that runs on any Windows 10/11 PC
rem without Python installed. Run this again after changing the code.
cd /d "%~dp0"
python -m pip install -q pyinstaller pdfplumber openpyxl
python -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name "Statement to Excel" ^
  --icon icon.ico --add-data "icon.ico;." ^
  --collect-data pdfminer ^
  --exclude-module numpy --exclude-module pandas --exclude-module scipy ^
  --exclude-module matplotlib --exclude-module IPython ^
  app.py
echo.
echo Done: dist\Statement to Excel.exe
pause
