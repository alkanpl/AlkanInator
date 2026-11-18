@echo off
rem AlkanInator: instalacja bibliotek do pobierania i wysylki plikow do sklepu (woo_api_cli.py, "Wgraj zmiany do sklepu").
cd /d "%~dp0"
echo === Python ===
py --version
if errorlevel 1 (
  echo.
  echo Brak Pythona - polecenie "py" nie dziala.
  echo Zainstaluj Python z https://www.python.org/downloads/
  echo i w instalatorze zostaw zaznaczone "py launcher".
  goto koniec
)
echo.
echo === pip ===
py -m pip --version
if errorlevel 1 py -m ensurepip --upgrade
echo.
echo === Instalacja bibliotek: requests, openpyxl, Pillow ===
py -m pip install requests openpyxl Pillow
echo.
echo === Sprawdzenie ===
py -c "import requests, openpyxl, PIL; print('Biblioteki OK')"
if errorlevel 1 (
  echo Instalacja nie powiodla sie - skopiuj komunikaty z tego okna.
  goto koniec
)
if not exist "src\woo_api\config.json" (
  echo.
  echo UWAGA: brak pliku src\woo_api\config.json - bez niego nie bedzie polaczenia ze sklepem.
  echo Token kazdy generuje sobie w WooCommerce, zakladka AlkanInator API.
)
echo.
echo Gotowe. Teraz uruchom "Wgraj zmiany do sklepu.bat".
:koniec
echo.
pause
