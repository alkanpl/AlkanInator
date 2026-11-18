@echo off
rem AlkanInator: podglad i zapis zmian z pliku XLSX do sklepu - otwiera sie w przegladarce.
cd /d "%~dp0"
rem Brak bibliotek (nowy komputer) - najpierw instalacja, zeby pyw nie padl bez komunikatu.
py -c "import requests, openpyxl, PIL" >nul 2>&1
if errorlevel 1 (
  call "Instalacja.bat"
  py -c "import requests, openpyxl, PIL" >nul 2>&1
  if errorlevel 1 exit /b 1
)
start "" pyw "src\woo_preview_app.py"
