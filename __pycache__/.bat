@echo off
cd /d "%~dp0"
set FLASK_APP=app.py
set FLASK_ENV=development
python -m flask run
pause
