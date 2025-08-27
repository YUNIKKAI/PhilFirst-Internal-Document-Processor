@echo off
echo 🛑 Stopping PhilFirst Renewal Notice Extractor
echo ================================================================

REM 1. Check if virtual environment exists
if not exist ".venv" (
    echo ❌ Virtual environment not found. Nothing to stop.
    pause
    exit /b 1
)

REM 2. Find and kill Gunicorn process
echo 🔎 Looking for Gunicorn process...
for /f "tokens=2 delims=," %%a in ('tasklist /FI "IMAGENAME eq gunicorn.exe" /FO CSV /NH') do (
    echo ⚡ Killing process ID %%a
    taskkill /PID %%a /F
    goto :done
)

echo ⚠️ No Gunicorn process found.
goto :end

:done
echo ✅ Gunicorn stopped successfully.

:end
pause
