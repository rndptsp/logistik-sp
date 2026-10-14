@echo off
cd /d "%~dp0.."
REM Builder master baru (MASTER BARU). Cadangan master lama: run_build_lama.bat
python tools\build_site_baru.py --master x %*
pause
