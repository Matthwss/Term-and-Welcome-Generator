@echo off
cd /d "%~dp0"
python gerar_termo.py
if errorlevel 1 pause
