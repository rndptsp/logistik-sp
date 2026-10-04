@echo off
cd /d "%~dp0.."
python tools\build_site.py %*
pause
