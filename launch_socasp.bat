@echo off
cd /d "%~dp0"
panel serve .\socasp.py --cookie-secret toor --port 80 --show --autoreload
pause
