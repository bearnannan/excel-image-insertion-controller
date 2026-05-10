@echo off
REM =================================================================
REM Project: Excel Image Insertion Controller
REM Author: WATCHARA MANADEE
REM Description: Build script for PyInstaller
REM =================================================================
echo Installing requirements...
pip install -r requirements.txt

echo Building Executable...
pyinstaller --noconsole --onefile --icon=icon.ico --add-data "icon.ico;." --collect-all customtkinter --collect-data tkinterdnd2 main.py

echo.
echo Build Complete! Check the 'dist' folder for main.exe
pause
