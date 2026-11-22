@echo off
REM ============================================================
REM  Eenmalig uitvoeren op EEN Windows-pc waar Python staat.
REM  Resultaat: dist\InzendingenSamenvoegen.exe
REM  Dat exe-bestand kunnen collega's downloaden en dubbelklikken;
REM  zij hebben geen Python nodig (wel Microsoft Word voor Word-inzendingen).
REM ============================================================
python -m pip install --upgrade pip
python -m pip install pypdf reportlab pywin32 pyinstaller
python -m PyInstaller --onefile --windowed --name InzendingenSamenvoegen ^
  --hidden-import win32timezone --hidden-import pythoncom --hidden-import win32com.client ^
  inzendingen_samenvoegen.py
echo.
echo Klaar. Het programma staat in de map "dist".
pause
