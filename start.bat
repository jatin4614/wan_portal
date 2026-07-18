@echo off
cd /d "%~dp0"

if not exist venv (
  echo Creating virtual environment...
  python -m venv venv
)

call venv\Scripts\activate.bat
pip install -q -r requirements.txt

REM Upload cap: set to 0 to disable, or any number of MB.
REM Defaults to 16 GB if unset.
if "%MAX_UPLOAD_MB%"=="" set MAX_UPLOAD_MB=16384

echo Starting Vajr Comn State on http://0.0.0.0:5000 (upload cap: %MAX_UPLOAD_MB% MB)
python app.py
pause
