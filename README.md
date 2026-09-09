# Infusion Reconciliation Pipeline

## Main Execution Script
The primary code script that executes the Controlled Substance Infusion Reconciliation process is:

- **Main Processor Script**: `step1_pyxis_processor.py` (invoked via `python src/main.py`)
- **Batch Entrypoint**: `run_reconciliation.bat`
- **PowerShell Entrypoint**: `run_reconciliation.ps1`

## How to Run
```bash
python src/main.py
```
or run `run_reconciliation.bat` / `run_reconciliation.ps1`.

## Input Directory
Input files are placed in `Input/` directory (e.g. Pyxis report, Epic Administrations, Epic Flowsheets).

## Output Files
The pipeline outputs reports to `Output/`:
- `Infusion_Reconciliation_Report_<YYYYMMDD_HHMMSS>.xlsx` (Timestamped report)
- `Infusion_Reconciliation_Report.xlsx` (Latest report)
- `Step1_Pyxis_Analysis_<date>.xlsx`
