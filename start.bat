@echo off
cd /d "%~dp0"
where py >nul 2>nul
if errorlevel 1 goto python
py -3 -m venv .venv
if errorlevel 1 goto fail
goto install
:python
python -m venv .venv
if errorlevel 1 goto fail
:install
.venv\Scripts\python -m pip install -r requirements.txt
if errorlevel 1 goto fail
.venv\Scripts\python app.py
pause
exit /b
:fail
echo Setup failed. Install Python 3.10+ and check your network.
pause
