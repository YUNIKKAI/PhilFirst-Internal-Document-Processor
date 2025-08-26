@echo off
echo 🚀 Starting PhilFirst Renewal Notice Extractor in Production Mode
echo ==================================================================

REM --- 1. Check if app.py exists
if not exist "app.py" (
    echo ❌ app.py not found. Please run from project directory.
    pause
    exit /b 1
)

REM --- 2. Create logs and static directories if missing
if not exist "logs" mkdir logs
if not exist "static" mkdir static

REM --- 3. Set environment variables
set FLASK_ENV=production

REM --- 4. Set default port if not provided
if "%1"=="" (
    set PORT=8000
) else (
    set PORT=%1
)

REM --- 5. Generate daily log filename (YYYY-MM-DD)
for /f "tokens=2 delims==" %%i in ('wmic os get localdatetime /value') do set ldt=%%i
set DATESTAMP=%ldt:~0,4%-%ldt:~4,2%-%ldt:~6,2%
set LOGFILE=logs\production_%DATESTAMP%.log

REM --- 6. Auto-clean logs older than 7 days
forfiles /p "logs" /m production_*.log /d -7 /c "cmd /c del @path" >nul 2>&1

REM --- 7. Start NGINX in WSL (optional)
echo 🌐 Starting NGINX in WSL...
wsl -d Ubuntu-22.04 -- sudo service nginx start

REM --- 8. Ensure Linux venv exists, create if missing
echo 🐍 Checking WSL venv...
wsl -d Ubuntu-22.04 -- bash -c "if [ ! -d 'venv' ]; then python3 -m venv venv && echo '✅ Linux venv created'; fi"

REM --- 9. Install requirements in Linux venv
echo 📦 Installing dependencies in WSL venv...
wsl -d Ubuntu-22.04 -- bash -c "source venv/bin/activate && pip install --upgrade pip && pip install -r requirements.txt"

REM --- 10. Start Gunicorn in WSL in new console
echo 🌟 Starting Gunicorn in WSL on port %PORT%, logging to %LOGFILE%...
start cmd /k wsl -d Ubuntu-22.04 -- bash -c "cd /mnt/c/Users/User/PhilFirst-Internal-Document-Processor && source venv/bin/activate && gunicorn -c gunicorn.conf.py -b 127.0.0.1:%PORT% wsgi:app >> /mnt/c/Users/User/PhilFirst-Internal-Document-Processor/%LOGFILE% 2>&1"

echo ✅ Production environment initialized.
pause
