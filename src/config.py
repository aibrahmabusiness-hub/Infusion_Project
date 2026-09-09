import os

# Root directory of the project
BASE_DIR = r"c:\Users\GaneshBhat\OneDrive - Novatio Solutions\Desktop\Infusion"

# Input paths
INPUT_DIR = os.path.join(BASE_DIR, "Input")
PYXIS_PATH = os.path.join(INPUT_DIR, "AuditTransactionDetail_RC.csv")
EPIC_ADMIN_PATH = os.path.join(INPUT_DIR, "BCH_Administered_Medications_Report_20260617_1354.xlsx")
EPIC_FLOWSHEET_PATH = os.path.join(INPUT_DIR, "BCH_Flowsheet_Med_Infusion_Volume_20260617_1342.xlsx")

# Output paths
OUTPUT_DIR = os.path.join(BASE_DIR, "Output")
os.makedirs(OUTPUT_DIR, exist_ok=True)
FINAL_REPORT_PATH = os.path.join(OUTPUT_DIR, "Infusion_Reconciliation_Report.xlsx")

# State tracking path
HISTORICAL_STATE_PATH = os.path.join(BASE_DIR, "historical_orders.json")

# Business Rules and Thresholds
MAX_HANG_TIME_HOURS = 96.0  # 96 hours maximum infusion hang time
VEND_WASTE_MATCH_WINDOW_HOURS = 8.0  # Time window for matching vends and wastes

# Match parameters
AI_MATCH_CONFIDENCE_THRESHOLD = 0.85

# Column mapping keys (CSV / Excel names)
# Pyxis Columns of Interest
PYXIS_COLS = {
    'order_number': 'OrderNumber',
    'patient_id': 'PatientID',
    'patient_name': 'PatientName',
    'tx_time': 'TransactionDateTime',
    'tx_type': 'TransactionType',
    'med_desc': 'MedDescription',
    'dispense_amt': 'DispenseAmount',
    'waste_amt': 'WasteAmount'
}

# Epic Admin Columns of Interest
EPIC_ADMIN_COLS = {
    'mrn': 'MRN',
    'patient_name': 'Patient',
    'order_id': 'Order ID',
    'medication': 'Medication',
    'action': 'AdministrationAction',
    'time': 'AdministrationInstant',
    'dose': 'Dose',
    'unit': 'DoseUnit'
}

# Epic Flowsheet Columns of Interest
EPIC_FLOWSHEET_COLS = {
    'mrn': 'MRN',
    'patient_name': 'Patient Name',
    'order_id': 'Order Med ID',
    'row_name': 'Flowsheet Row Name',
    'display_name': 'Flowsheet Display',
    'time': 'Recorded Time',
    'value': 'Flowsheet Value'
}
