@echo off
chcp 65001 >nul
cd /d "%~dp0"
py -3 -m pip install --upgrade pyinstaller
pyinstaller --noconfirm --clean --name SeatRotator --console --paths app --add-data "app/web;app/web" windows_entry.py
pause
