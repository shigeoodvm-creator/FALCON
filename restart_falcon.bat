@echo off
echo Restarting FALCON...
taskkill /IM python.exe /F >nul 2>&1
timeout /t 1 >nul
python main.py
