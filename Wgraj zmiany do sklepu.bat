@echo off
rem AlkanInator: podglad i zapis zmian z pliku XLSX do sklepu - otwiera sie w przegladarce.
cd /d "%~dp0"
start "" pyw "src\woo_preview_app.py"
