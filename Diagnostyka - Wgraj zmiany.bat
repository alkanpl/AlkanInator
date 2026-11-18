@echo off
rem AlkanInator: diagnostyka aplikacji "Wgraj zmiany do sklepu" - okno zostaje otwarte z bledami.
cd /d "%~dp0"
echo === Python ===
py --version
echo.
echo === Biblioteki (requests, openpyxl, Pillow) ===
py -c "import requests, openpyxl, PIL; print('OK')"
echo.
echo === Start aplikacji (zamkniecie okna zatrzymuje aplikacje) ===
py src\woo_preview_app.py
echo.
pause
