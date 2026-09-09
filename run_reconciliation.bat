@echo off
echo ============================================================
echo Installing dependencies from requirements.txt...
echo ============================================================
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install dependencies.
    pause
    exit /b 1
)

echo ============================================================
echo Running Reconciliation Pipeline...
echo ============================================================
python src\main.py
if %errorlevel% neq 0 (
    echo [ERROR] Reconciliation pipeline failed.
    pause
    exit /b 1
)

echo ============================================================
echo Reconciliation complete!
echo ============================================================
pause
