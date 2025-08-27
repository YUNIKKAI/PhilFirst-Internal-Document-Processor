@echo off
echo 🚀 Deploying PhilFirst Renewal Notice Extractor
echo ================================================================

REM 1. Delete old .venv if it exists
if exist .venv (
    echo 🗑️ Removing old virtual environment (.venv)...
    rmdir /s /q .venv
)

REM 2. Create new .venv
echo 🐍 Creating new virtual environment...
python -m venv .venv

REM 3. Activate .venv
call .venv\Scripts\activate.bat

REM 4. Upgrade pip
echo 🔄 Upgrading pip...
python -m pip install --upgrade pip

REM 5. Install requirements
if exist requirements.txt (
    echo 📦 Installing dependencies from requirements.txt...
    pip install -r requirements.txt
) else (
    echo ⚠️ requirements.txt not found, skipping dependency install.
)

REM 6. Verify script exists
if not exist verify_packages.py (
    echo ⚠️ verify_packages.py not found. Please add it manually.
) else (
    echo 🛠️ Using existing verify_packages.py
)

echo ✅ Deployment complete!
pause
