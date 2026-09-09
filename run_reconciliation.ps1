Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "Installing dependencies from requirements.txt..." -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
pip install -r requirements.txt
if ($LASTEXITCODE -ne 0) {
    Write-Error "Failed to install dependencies."
    Read-Host "Press Enter to exit..."
    exit 1
}

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "Running Reconciliation Pipeline..." -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
python src\main.py
if ($LASTEXITCODE -ne 0) {
    Write-Error "Reconciliation pipeline failed."
    Read-Host "Press Enter to exit..."
    exit 1
}

Write-Host "============================================================" -ForegroundColor Green
Write-Host "Reconciliation complete!" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Read-Host "Press Enter to exit..."
