import os
import re
import math
import json
import pandas as pd
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# --- CONFIGURATIONS ---
BASE_DIR = r"c:\Users\GaneshBhat\OneDrive - Novatio Solutions\Desktop\Infusion"
PYXIS_CSV_PATH = os.path.join(BASE_DIR, "Input", "AuditTransactionDetail_RC.csv")
EPIC_ADMIN_PATH = os.path.join(BASE_DIR, "Input", "BCH_Administered_Medications_Report_20260617_1354.xlsx")
EPIC_FLOW_PATH = os.path.join(BASE_DIR, "Input", "BCH_Flowsheet_Med_Infusion_Volume_20260617_1342.xlsx")
OUTPUT_PATH = os.path.join(BASE_DIR, "Output", "Step1_Pyxis_Analysis.xlsx")
os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)

# Styles
NAVY_HEADER_FILL = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
NAVY_HEADER_FONT = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
SECTION_HEADER_FILL = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
SECTION_HEADER_FONT = Font(name="Segoe UI", size=12, bold=True, color="1A202C")

GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
GREEN_FONT = Font(name="Segoe UI", size=10, color="006100", bold=True)
RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
RED_FONT = Font(name="Segoe UI", size=10, color="9C0006", bold=True)
YELLOW_FILL = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
YELLOW_FONT = Font(name="Segoe UI", size=10, color="9C6500", bold=True)

ORANGE_FILL = PatternFill(start_color="FFE8D6", end_color="FFE8D6", fill_type="solid")
ORANGE_FONT = Font(name="Segoe UI", size=10, color="A75D00", bold=True)

ZEBRA_FILL = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

TITLE_FONT = Font(name="Segoe UI", size=16, bold=True, color="1B365D")
SUBTITLE_FONT = Font(name="Segoe UI", size=11, italic=True, color="4A5568")
LABEL_FONT = Font(name="Segoe UI", size=10, bold=True, color="4A5568")
DATA_FONT = Font(name="Segoe UI", size=10, color="000000")
BOLD_DATA_FONT = Font(name="Segoe UI", size=10, bold=True, color="000000")

THIN_BORDER_SIDE = Side(border_style="thin", color="CBD5E1")
THIN_BORDER = Border(left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=THIN_BORDER_SIDE, bottom=THIN_BORDER_SIDE)
KPI_BORDER = Border(left=Side(border_style="medium", color="1B365D"), right=Side(border_style="medium", color="1B365D"), top=Side(border_style="medium", color="1B365D"), bottom=Side(border_style="medium", color="1B365D"))

# --- UTILITIES ---
def parse_dates_from_filename(filename):
    match = re.search(r"(\d{8})[_\-](\d{8})", filename)
    if match:
        try:
            start_dt = datetime.strptime(match.group(1), "%Y%m%d")
            end_dt = datetime.strptime(match.group(2), "%Y%m%d")
            return start_dt, end_dt
        except Exception:
            pass
    return None, None

def get_order_window(p_sub, e_prev, e_this, time_window_hours_before=1.0, max_hang_time_hours=96.0):
    if p_sub.empty:
        return pd.NaT, pd.NaT
        
    p_times = sorted(p_sub['TransactionDateTime_Clean'].dropna().tolist())
    if not p_times:
        return pd.NaT, pd.NaT
        
    first_tx_time = p_times[0]
    last_tx_time = p_times[-1]
    
    first_tx_type = p_sub.sort_values('TransactionDateTime_Clean').iloc[0]['TransactionType_Clean']
    
    if first_tx_type == 'Waste':
        e_all = pd.concat([e_prev, e_this]) if (not e_prev.empty or not e_this.empty) else pd.DataFrame()
        t_dispense = None
        if not e_all.empty:
            e_before = e_all[e_all['AdministrationInstant_Clean'] < first_tx_time].sort_values('AdministrationInstant_Clean', ascending=False)
            if not e_before.empty:
                for _, r in e_before.iterrows():
                    act_lower = str(r['AdministrationAction']).lower()
                    if any(kw in act_lower for kw in ['bag', 'dose', 'start', 'restart']):
                        t_dispense = r['AdministrationInstant_Clean']
                        break
        if t_dispense:
            start_w = t_dispense - pd.Timedelta(hours=time_window_hours_before)
        else:
            start_w = first_tx_time - pd.Timedelta(hours=max_hang_time_hours)
    else:
        start_w = first_tx_time - pd.Timedelta(hours=time_window_hours_before)
        
    end_w = last_tx_time + pd.Timedelta(hours=max_hang_time_hours)
    return start_w, end_w


def clean_id(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s

def clean_patient_name(name):
    if pd.isna(name) or not isinstance(name, str):
        return ""
    return "".join(c for c in name.lower() if c.isalnum())

def parse_amount_string(amount_str):
    if not isinstance(amount_str, str):
        if isinstance(amount_str, (int, float)) and not math.isnan(amount_str):
            return float(amount_str), None
        return None, None
    amount_str = amount_str.strip()
    match = re.match(r"^(\d+(?:\.\d+)?)\s*(mg|mcg|mCg|g|mL|ml|ML)$", amount_str, re.IGNORECASE)
    if match:
        val = float(match.group(1))
        unit = match.group(2).lower()
        return val, unit
    return None, None

def parse_med_description(desc):
    if not isinstance(desc, str):
        return None
    
    desc_lower = desc.lower()
    drug_name = "unknown"
    for drug in ['ketamine', 'midazolam', 'versed', 'fentanyl', 'morphine', 'hydromorphone', 'dilaudid']:
        if drug in desc_lower:
            drug_name = drug
            break
        
    m1 = re.search(r"(\d+(?:\.\d+)?)\s*(mg|mcg|mCg)/(\d+(?:\.\d+)?)\s*(mL|ml|ML)\s*\((\d+(?:\.\d+)?)\s*(mL|ml|ML)\)", desc, re.IGNORECASE)
    if m1:
        conc_val = float(m1.group(1))
        conc_unit = m1.group(2).lower()
        conc_vol = float(m1.group(3))
        tot_vol = float(m1.group(5))
        concentration = conc_val / conc_vol
        total_med = concentration * tot_vol
        return {
            'drug_name': drug_name,
            'concentration_strength': conc_val,
            'concentration_unit': conc_unit,
            'concentration_volume': conc_vol,
            'total_volume': tot_vol,
            'concentration_per_ml': concentration,
            'calculated_total_amount': total_med,
            'calculated_total_unit': conc_unit
        }
        
    m2 = re.search(r"(\d+(?:\.\d+)?)\s*(mg|mcg|mCg)\s*\((\d+(?:\.\d+)?)\s*(mL|ml|ML)\)", desc, re.IGNORECASE)
    if m2:
        tot_val = float(m2.group(1))
        tot_unit = m2.group(2).lower()
        tot_vol = float(m2.group(3))
        concentration = tot_val / tot_vol
        return {
            'drug_name': drug_name,
            'concentration_strength': tot_val,
            'concentration_unit': tot_unit,
            'concentration_volume': tot_vol,
            'total_volume': tot_vol,
            'concentration_per_ml': concentration,
            'calculated_total_amount': tot_val,
            'calculated_total_unit': tot_unit
        }
        
    return {
        'drug_name': drug_name,
        'concentration_strength': None,
        'concentration_unit': None,
        'concentration_volume': None,
        'total_volume': None,
        'concentration_per_ml': None,
        'calculated_total_amount': None,
        'calculated_total_unit': None
    }

def convert_mass_to_volume(mass_val, mass_unit, conc_per_ml, conc_unit):
    if mass_val is None or conc_per_ml is None or conc_per_ml == 0:
        return None
    m_unit = mass_unit.lower() if isinstance(mass_unit, str) else ""
    c_unit = conc_unit.lower() if isinstance(conc_unit, str) else ""
    factor = 1.0
    if m_unit == 'mcg' and c_unit == 'mg':
        factor = 0.001
    elif m_unit == 'mg' and c_unit == 'mcg':
        factor = 1000.0
    return (mass_val * factor) / conc_per_ml

def match_medication_names(med_pyxis, med_epic):
    if pd.isna(med_pyxis) or pd.isna(med_epic):
        return False
    p_lower = str(med_pyxis).lower()
    e_lower = str(med_epic).lower()
    
    synonyms = [
        ('versed', 'midazolam'),
        ('dilaudid', 'hydromorphone')
    ]
    
    for drug in ['fentanyl', 'midazolam', 'versed', 'ketamine', 'morphine', 'hydromorphone', 'dilaudid']:
        if drug in p_lower and drug in e_lower:
            return True
        for syn1, syn2 in synonyms:
            if (syn1 in p_lower or syn2 in p_lower) and (syn1 in e_lower or syn2 in e_lower):
                return True
    return False

# Step 3: Flowsheet Drug synonym logic (100% confidence)
def match_flowsheet_med(pyxis_med_desc, flowsheet_row, flowsheet_disp):
    p_med = str(pyxis_med_desc).lower()
    f_row = str(flowsheet_row).lower()
    f_disp = str(flowsheet_disp).lower()
    
    drug_map = {
        'ketamine': ['ketamine'],
        'midazolam': ['midazolam', 'versed'],
        'fentanyl': ['fentanyl'],
        'morphine': ['morphine'],
        'hydromorphone': ['hydromorphone', 'dilaudid']
    }
    
    active_drug = None
    for k, syns in drug_map.items():
        if any(syn in p_med for syn in syns):
            active_drug = k
            break
            
    if not active_drug:
        return False, 0.0
        
    if active_drug in f_row or active_drug in f_disp:
        return True, 1.0
    if active_drug == 'midazolam' and ('versed' in f_row or 'versed' in f_disp):
        return True, 1.0
    if active_drug == 'hydromorphone' and ('dilaudid' in f_row or 'dilaudid' in f_disp):
        return True, 1.0
        
    return False, 0.0

def load_excel_safe(path):
    if not path or not os.path.exists(path):
        return pd.DataFrame()
    try:
        if str(path).lower().endswith('.csv'):
            return pd.read_csv(path, low_memory=False)
        return pd.read_excel(path)
    except PermissionError:
        print(f"File locked: {path}. Copying to read safely...")
        temp_path = "temp_" + os.path.basename(path)
        import shutil
        try:
            shutil.copy(path, temp_path)
        except Exception:
            import subprocess
            subprocess.run(['powershell', '-Command', f'Copy-Item -Path "{path}" -Destination "{temp_path}" -Force'], shell=True)
        if str(path).lower().endswith('.csv'):
            df = pd.read_csv(temp_path, low_memory=False)
        else:
            df = pd.read_excel(temp_path)
        try:
            os.remove(temp_path)
        except:
            pass
        return df

def normalize_flowsheet_columns(df):
    if df.empty:
        return df
    col_map = {
        'ORDER_MED_ID': 'Order Med ID',
        'PAT_NAME': 'Patient Name',
        'RECORDED_TIME': 'Recorded Time',
        'MEAS_VALUE': 'Flowsheet Value',
        'FLO_MEAS_NAME': 'Flowsheet Row Name',
        'DISP_NAME': 'Flowsheet Display'
    }
    rename_dict = {}
    for col in df.columns:
        c_clean = str(col).strip()
        if c_clean in col_map:
            rename_dict[col] = col_map[c_clean]
        elif c_clean.upper() in col_map:
            rename_dict[col] = col_map[c_clean.upper()]
    return df.rename(columns=rename_dict)

def normalize_admin_columns(df):
    if df.empty:
        return df
    col_map = {
        'ORDER_ID': 'Order ID',
        'PATIENT_NAME': 'Patient',
        'PAT_NAME': 'Patient',
        'ADMINISTRATION_ACTION': 'AdministrationAction',
        'ADMIN_ACTION': 'AdministrationAction',
        'ADMINISTRATION_INSTANT': 'AdministrationInstant',
        'ADMIN_INSTANT': 'AdministrationInstant',
        'DOSE_UNIT': 'DoseUnit'
    }
    rename_dict = {}
    for col in df.columns:
        c_clean = str(col).strip()
        if c_clean in col_map:
            rename_dict[col] = col_map[c_clean]
        elif c_clean.upper() in col_map:
            rename_dict[col] = col_map[c_clean.upper()]
    return df.rename(columns=rename_dict)

def clean_val_for_json(val):
    if val is None or pd.isna(val) or math.isnan(val) or math.isinf(val):
        return 0.0
    return float(val)

def make_pyxis_signature(row):
    fields = [
        str(row.get('StationName', '')).strip(),
        str(row.get('TransactionDateTime', '')).strip(),
        str(row.get('UserName', '')).strip(),
        str(row.get('OrderNumber', '')).strip(),
        str(row.get('TransactionType', '')).strip(),
        str(row.get('DispenseAmount', '')).strip(),
        str(row.get('WasteAmount', '')).strip()
    ]
    return "||".join(fields)

def make_flowsheet_signature(row):
    fields = [
        str(row.get('CSN', '')).strip(),
        str(row.get('Flowsheet Row Name', '')).strip(),
        str(row.get('Recorded Time', '')).strip(),
        str(row.get('Flowsheet Value', '')).strip()
    ]
    return "||".join(fields)

def determine_status(tx_type):
    if tx_type == 'Vend':
        return 'Active'
    elif tx_type == 'Waste':
        return 'Active'
    elif tx_type == 'Stock Return':
        return 'Stock Return'
    elif tx_type == 'Vend (Cancel)':
        return 'Cancelled'
    return 'Unknown'

def clean_pyxis_df(df):
    df = df.copy()
    df['OrderNumber_Clean'] = df['OrderNumber'].apply(clean_id)
    df['PatientID_Clean'] = df['PatientID'].apply(clean_id)
    df['PatientName_Clean'] = df['PatientName'].apply(lambda x: str(x).strip() if pd.notna(x) else "Unknown")
    df['PatientName_Norm'] = df['PatientName'].apply(clean_patient_name)
    df['TransactionDateTime_Clean'] = pd.to_datetime(df['TransactionDateTime'], format='mixed', errors='coerce')
    df['TransactionType_Clean'] = df['TransactionType'].astype(str).str.strip()
    df_filt = df[df['TransactionType_Clean'].isin(['Vend', 'Waste', 'Stock Return', 'Vend (Cancel)'])].copy()
    df_filt['TransactionStatus'] = df_filt['TransactionType_Clean'].apply(determine_status)
    return df_filt

# --- MAIN RUN ---
def process_pyxis_file(pyxis_path, prev_pyxis_path=None):
    import glob
    filename = os.path.basename(pyxis_path)
    match = re.search(r"weekly report\s+([\d\-]+\s+[\d\-]+\s+[APap][Mm]|[\d\-]+)", filename, re.IGNORECASE)
    if match:
        date_str = match.group(1).replace(" ", "_")
    else:
        date_str = "Analysis"
    file_out = f"Step1_Pyxis_Analysis_{date_str}.xlsx"
    OUTPUT_PATH = os.path.join(BASE_DIR, "Output", file_out)
    
    print("\n" + "=" * 60)
    print(f"PROCESSING PYXIS FILE: {filename}")
    print(f"Output Path: {OUTPUT_PATH}")
    print("=" * 60)
    
    df_pyxis_raw = pd.read_csv(pyxis_path, low_memory=False)
    df_pyxis_raw['TransactionDateTime_Clean'] = pd.to_datetime(df_pyxis_raw['TransactionDateTime'], format='mixed', errors='coerce')
    
    pyxis_min_time = df_pyxis_raw['TransactionDateTime_Clean'].min()
    pyxis_max_time = df_pyxis_raw['TransactionDateTime_Clean'].max()
    print(f"Pyxis date range in file: {pyxis_min_time} to {pyxis_max_time}")
    
    # 2. Locate corresponding This Week and Previous Week Epic files
    input_dir = os.path.dirname(pyxis_path)
    all_input_files = [os.path.join(input_dir, f) for f in os.listdir(input_dir)
                       if f.lower().endswith(('.xlsx', '.xls', '.csv')) and os.path.abspath(os.path.join(input_dir, f)) != os.path.abspath(pyxis_path)]
    
    this_week_flow = None
    prev_week_flow = None
    this_week_admin = None
    prev_week_admin = None
    
    for f in all_input_files:
        name = os.path.basename(f)
        if "cs continuous infusions" in name.lower() or "audittransaction" in name.lower():
            continue
        start_dt, end_dt = parse_dates_from_filename(name)
        if start_dt and end_dt:
            if start_dt.date() <= pyxis_min_time.date() <= end_dt.date():
                if "flowsheet" in name.lower() or "volume" in name.lower():
                    this_week_flow = f
                elif "admin" in name.lower() or "action" in name.lower():
                    this_week_admin = f
            elif end_dt.date() < pyxis_min_time.date() and (pyxis_min_time.date() - end_dt.date()).days <= 7:
                if "flowsheet" in name.lower() or "volume" in name.lower():
                    prev_week_flow = f
                elif "admin" in name.lower() or "action" in name.lower():
                    prev_week_admin = f
                    
    print(f"Detected This Week Flowsheet: {os.path.basename(this_week_flow) if this_week_flow else 'Not Found'}")
    print(f"Detected Previous Week Flowsheet: {os.path.basename(prev_week_flow) if prev_week_flow else 'Not Found'}")
    print(f"Detected This Week Admin: {os.path.basename(this_week_admin) if this_week_admin else 'Not Found'}")
    print(f"Detected Previous Week Admin: {os.path.basename(prev_week_admin) if prev_week_admin else 'Not Found'}")
    
    # 3. Load Epic Files
    df_flow_this = normalize_flowsheet_columns(load_excel_safe(this_week_flow)) if this_week_flow else pd.DataFrame()
    df_flow_prev = normalize_flowsheet_columns(load_excel_safe(prev_week_flow)) if prev_week_flow else pd.DataFrame()
    df_epic_this = normalize_admin_columns(load_excel_safe(this_week_admin)) if this_week_admin else pd.DataFrame()
    df_epic_prev = normalize_admin_columns(load_excel_safe(prev_week_admin)) if prev_week_admin else pd.DataFrame()
    
    # Load Previous Week Pyxis if available
    df_pyxis_prev = pd.DataFrame()
    if prev_pyxis_path and os.path.exists(prev_pyxis_path):
        print(f"Detected Previous Week Pyxis: {os.path.basename(prev_pyxis_path)}")
        try:
            df_pyxis_prev_raw = pd.read_csv(prev_pyxis_path, low_memory=False)
            df_pyxis_prev = clean_pyxis_df(df_pyxis_prev_raw)
        except Exception as e:
            print(f"Warning: Failed to load previous week Pyxis data: {e}")
            
    # Preprocess Pyxis
    df_pyxis_filt = clean_pyxis_df(df_pyxis_raw)
    
    # Preprocess Epic Administrations (This Week & Prev Week)
    if not df_epic_this.empty:
        df_epic_this['Audit Week'] = 'Current Week'
    if not df_epic_prev.empty:
        df_epic_prev['Audit Week'] = 'Previous Week'
    for df in [df_epic_this, df_epic_prev]:
        if not df.empty:
            df['OrderID_Clean'] = df['Order ID'].apply(clean_id)
            df['MRN_Clean'] = df['MRN'].apply(clean_id)
            df['PatientName_Clean'] = df['Patient'].astype(str).str.strip()
            df['PatientName_Norm'] = df['Patient'].apply(clean_patient_name)
            df['AdministrationInstant_Clean'] = pd.to_datetime(df['AdministrationInstant'], format='mixed', errors='coerce')
            
    # Preprocess Epic Flowsheets (This Week & Prev Week)
    if not df_flow_this.empty:
        df_flow_this['Audit Week'] = 'Current Week'
    if not df_flow_prev.empty:
        df_flow_prev['Audit Week'] = 'Previous Week'
    for df in [df_flow_this, df_flow_prev]:
        if not df.empty:
            df['OrderID_Clean'] = df['Order Med ID'].apply(clean_id)
            df['MRN_Clean'] = df['MRN'].apply(clean_id)
            df['PatientName_Norm'] = df['Patient Name'].apply(clean_patient_name)
            df['RecordedTime_Clean'] = pd.to_datetime(df['Recorded Time'], format='mixed', errors='coerce')
            df['FlowsheetValue_Num'] = pd.to_numeric(df['Flowsheet Value'], errors='coerce').fillna(0.0)
            
    # Define unified dataframes for logging compatibility
    df_epic = pd.concat([df_epic_prev, df_epic_this], ignore_index=True).drop_duplicates() if (not df_epic_prev.empty or not df_epic_this.empty) else pd.DataFrame()
    df_flow = pd.concat([df_flow_prev, df_flow_this], ignore_index=True).drop_duplicates() if (not df_flow_prev.empty or not df_flow_this.empty) else pd.DataFrame()

    # Track flowsheet and administration log rows to output to extra sheets
    prev_week_flows_collected = []
    this_week_flows_collected = []
    prev_week_admins_collected = []
    this_week_admins_collected = []
    prev_pyxis_rows = []
            
    # Drive order matching strictly from Pyxis
    all_order_ids = set(df_pyxis_filt['OrderNumber_Clean'].unique()) - {''}
    
    matching_rows = []
    exceptions = []
    
    # Dictionary to map Order ID -> Reconciliation Unit & Details
    order_recon_info = {}
    
    # Track flowsheet log rows to output later
    flowsheet_log_rows = []
    
    for order_id in all_order_ids:
        p_sub = df_pyxis_filt[df_pyxis_filt['OrderNumber_Clean'] == order_id]
        e_this = df_epic_this[df_epic_this['OrderID_Clean'] == order_id] if not df_epic_this.empty else pd.DataFrame()
        e_prev = df_epic_prev[df_epic_prev['OrderID_Clean'] == order_id] if not df_epic_prev.empty else pd.DataFrame()
        f_this = df_flow_this[df_flow_this['OrderID_Clean'] == order_id] if not df_flow_this.empty else pd.DataFrame()
        f_prev = df_flow_prev[df_flow_prev['OrderID_Clean'] == order_id] if not df_flow_prev.empty else pd.DataFrame()
        
        # Combine all previous and current week records for calculations (no window filtering)
        f_this_filt = f_this
        f_prev_filt = f_prev
        e_this_filt = e_this
        e_prev_filt = e_prev
        e_sub = pd.concat([e_prev, e_this], ignore_index=True)
        f_sub = pd.concat([f_prev, f_this], ignore_index=True)
        
        pyxis_patient = p_sub.iloc[0]['PatientName'] if not p_sub.empty else "N/A"
        epic_patient = e_sub.iloc[0]['PatientName_Clean'] if not e_sub.empty else "N/A"
        flow_patient = f_sub.iloc[0]['Patient Name'] if not f_sub.empty else "N/A"
        
        pyxis_med = p_sub.iloc[0]['MedDescription'] if not p_sub.empty else "N/A"
        epic_med = e_sub.iloc[0]['Medication'] if not e_sub.empty else "N/A"
        
        # Mapped Flowsheet drug name
        flowsheet_drug_row = None
        if pyxis_med != "N/A":
            desc_lower = pyxis_med.lower()
            if 'ketamine' in desc_lower:
                flowsheet_drug_row = 'R KETAMINE VOLUME'
            elif 'midazolam' in desc_lower or 'versed' in desc_lower:
                flowsheet_drug_row = 'R MIDAZOLAM VOLUME'
            elif 'fentanyl' in desc_lower:
                flowsheet_drug_row = 'R FENTANYL VOLUME'
            elif 'morphine' in desc_lower:
                flowsheet_drug_row = 'R MORPHINE VOLUME'
            elif 'hydromorphone' in desc_lower or 'dilaudid' in desc_lower:
                flowsheet_drug_row = 'R HYDROMORPHONE VOLUME'
                
        # Collect flowsheet rows that are considered in calculations
        if flowsheet_drug_row:
            if not f_this_filt.empty:
                f_this_matched = f_this_filt[f_this_filt['Flowsheet Row Name'] == flowsheet_drug_row]
                if not f_this_matched.empty:
                    this_week_flows_collected.append(f_this_matched)
            if not f_prev_filt.empty:
                f_prev_matched = f_prev_filt[f_prev_filt['Flowsheet Row Name'] == flowsheet_drug_row]
                if not f_prev_matched.empty:
                    prev_week_flows_collected.append(f_prev_matched)
                    
        # Collect Epic Administrations that are considered for this order
        if not e_this_filt.empty:
            this_week_admins_collected.append(e_this_filt)
        if not e_prev_filt.empty:
            prev_week_admins_collected.append(e_prev_filt)
        
        is_matched = True
        match_details = []
        
        # Mismatch checks
        if e_sub.empty and f_sub.empty:
            is_matched = False
            match_details.append("Missing Epic administration and flowsheet records")
            exceptions.append({
                'OrderID': order_id,
                'PatientName': pyxis_patient,
                'Source': 'Epic Administrations',
                'ExceptionType': 'Missing Charting Activity',
                'Reason': f"Pyxis Order ID {order_id} has cabinet transactions but no administrations or flowsheet records charted in Epic.",
                'ActionNeeded': 'Verify if clinician omitted charting administration or if medication was wasted completely.'
            })
        else:
            # Check names in Administrations
            if not e_sub.empty:
                p_name_norm = p_sub.iloc[0]['PatientName_Norm']
                e_name_norm = e_sub.iloc[0]['PatientName_Norm']
                if p_name_norm != e_name_norm:
                    is_matched = False
                    match_details.append(f"Patient mismatch: Pyxis='{pyxis_patient}', Epic='{epic_patient}'")
                    exceptions.append({
                        'OrderID': order_id,
                        'PatientName': f"Pyxis: {pyxis_patient} | Epic: {epic_patient}",
                        'Source': 'Reconciliation Logic',
                        'ExceptionType': 'Patient ID Mismatch',
                        'Reason': f"Order {order_id} patient identities do not align between Pyxis and Epic Administrations.",
                        'ActionNeeded': 'Investigate correct patient record; check for duplicate Order IDs or spelling errors.'
                    })
                if not match_medication_names(pyxis_med, epic_med):
                    is_matched = False
                    match_details.append(f"Medication mismatch: Pyxis='{pyxis_med}', Epic='{epic_med}'")
                    exceptions.append({
                        'OrderID': order_id,
                        'PatientName': pyxis_patient,
                        'Source': 'Reconciliation Logic',
                        'ExceptionType': 'Medication Naming Mismatch',
                        'Reason': f"Medication names do not align between Pyxis ('{pyxis_med}') and Epic Administrations ('{epic_med}').",
                        'ActionNeeded': 'Confirm pharmacy formulary settings; verify if generic substitution was documented.'
                    })
            # Check names in flowsheets
            if not f_sub.empty:
                p_name_norm = p_sub.iloc[0]['PatientName_Norm']
                f_name_norm = f_sub.iloc[0]['PatientName_Norm']
                if p_name_norm != f_name_norm:
                    is_matched = False
                    match_details.append(f"Patient flowsheet mismatch: Pyxis='{pyxis_patient}', Flowsheet='{flow_patient}'")
                    exceptions.append({
                        'OrderID': order_id,
                        'PatientName': f"Pyxis: {pyxis_patient} | Flowsheet: {flow_patient}",
                        'Source': 'Epic Flowsheets',
                        'ExceptionType': 'Patient ID Mismatch',
                        'Reason': f"Order {order_id} patient identities do not align between Pyxis and Epic Flowsheets.",
                        'ActionNeeded': 'Check for duplicate Order IDs or patient profiling errors in flowsheet.'
                    })
                    
        # Unit and Action Rule Checking
        recon_unit = "mL"
        has_bolus = False
        stopped_restarted_actions = []
        
        if not e_sub.empty:
            actions = e_sub['AdministrationAction'].astype(str).str.strip().tolist()
            for action in actions:
                action_lower = action.lower()
                if 'bolus' in action_lower:
                    has_bolus = True
                if 'stop' in action_lower:
                    stopped_restarted_actions.append("Stopped")
                if 'restart' in action_lower:
                    stopped_restarted_actions.append("Restarted")
                    
            if has_bolus:
                parsed = parse_med_description(pyxis_med)
                recon_unit = parsed['concentration_unit'] if (parsed and parsed['concentration_unit']) else "mg"
            else:
                recon_unit = "mL"
        else:
            # Default Pyxis to mL if missing Epic context
            recon_unit = "mL"
            
        if stopped_restarted_actions:
            unique_actions = list(set(stopped_restarted_actions))
            exceptions.append({
                'OrderID': order_id,
                'PatientName': epic_patient if epic_patient != "N/A" else pyxis_patient,
                'Source': 'Epic Administrations',
                'ExceptionType': 'Stopped/Restarted Action Exception',
                'Reason': f"Order contains {', '.join(unique_actions)} actions. Manual pharmacy volume review is required.",
                'ActionNeeded': 'Audit pump flowsheet volumes around stopped/restarted events to check for correct waste documentation.'
            })
            
        order_recon_info[order_id] = {
            'unit': recon_unit,
            'has_bolus': has_bolus
        }
        
        # Get unique list of charted actions
        if not e_sub.empty:
            seen_act = set()
            unique_actions = [a for a in e_sub['AdministrationAction'].astype(str).str.strip().tolist() if not (a in seen_act or seen_act.add(a))]
            actions_str = ", ".join(unique_actions)
        else:
            actions_str = "N/A"
            
        match_status = "Matched" if is_matched else "Mismatch"
        details_str = ", ".join(match_details) if match_details else "Valid cross-system match"
        
        matching_rows.append({
            'OrderID': order_id,
            'PatientName_Pyxis': pyxis_patient,
            'PatientName_Epic': epic_patient,
            'Medication_Pyxis': pyxis_med,
            'Medication_Epic': epic_med,
            'MatchStatus': match_status,
            'ReconciliationUnit': recon_unit,
            'MatchDetails': details_str,
            'AdministrationActions': actions_str
        })
    
    # df_matching_overview creation deferred to line 887
    
    # 3. Process Pyxis with Unit Conversions & Pairing
    # Columns in df_pyxis_filt will hold converted values depending on order's reconciliation unit
    df_pyxis_filt['PairingStatus'] = 'Unmatched'
    df_pyxis_filt.loc[df_pyxis_filt['OrderNumber_Clean'] == '', 'PairingStatus'] = 'Unmatched - Missing Order ID'
    df_pyxis_filt.loc[df_pyxis_filt['TransactionStatus'].isin(['Stock Return', 'Cancelled']), 'PairingStatus'] = 'Excluded'
    df_pyxis_filt['PairedTxTime'] = None
    df_pyxis_filt['PairedTxType'] = None
    df_pyxis_filt['PairedTxAmount'] = None
    df_pyxis_filt['PairingGapHours'] = None
    df_pyxis_filt['ParsedConc'] = None
    df_pyxis_filt['ParsedConcUnit'] = None
    df_pyxis_filt['ParsedTotalVol'] = None
    df_pyxis_filt['CalculatedAmt_RecUnit'] = None
    df_pyxis_filt['RecUnit'] = None
    df_pyxis_filt['CalculatedDose'] = None
    df_pyxis_filt['CalculatedVol_ml'] = None
    df_pyxis_filt['TxIndex'] = range(len(df_pyxis_filt))
    
    df_pyxis_filt = df_pyxis_filt.sort_values('TransactionDateTime_Clean')
    
    paired_bags = []
    
    # Track order parsed concentrations for flowsheet conversion
    order_concentrations = {}
    
    for (order_num, med_desc), group in df_pyxis_filt.groupby(['OrderNumber_Clean', 'MedDescription']):
        if not order_num or pd.isna(med_desc):
            continue
            
        parsed = parse_med_description(med_desc)
        conc_per_ml = parsed['concentration_per_ml'] if parsed else None
        conc_unit = parsed['concentration_unit'] if parsed else None
        tot_vol = parsed['total_volume'] if parsed else None
        conc_str = f"{parsed['concentration_strength']} {parsed['concentration_unit']}/{parsed['concentration_volume']} mL" if (parsed and parsed['concentration_strength']) else "Unknown"
        
        order_concentrations[order_num] = {
            'conc_per_ml': conc_per_ml,
            'conc_unit': conc_unit
        }
        
        # Get reconciliation info for this Order ID
        r_info = order_recon_info.get(order_num, {'unit': 'mL', 'has_bolus': False})
        order_unit = r_info['unit']
        
        # Write parsed values to the group records
        for idx in group.index:
            df_pyxis_filt.loc[idx, 'ParsedConc'] = conc_str
            df_pyxis_filt.loc[idx, 'ParsedConcUnit'] = conc_unit
            df_pyxis_filt.loc[idx, 'ParsedTotalVol'] = tot_vol
            df_pyxis_filt.loc[idx, 'RecUnit'] = order_unit
            
            tx_type = df_pyxis_filt.loc[idx, 'TransactionType_Clean']
            raw_amt = df_pyxis_filt.loc[idx, 'DispenseAmount'] if tx_type == 'Vend' else df_pyxis_filt.loc[idx, 'WasteAmount']
            
            val, unit = parse_amount_string(raw_amt)
            if val is not None:
                df_pyxis_filt.loc[idx, 'CalculatedDose'] = val
                df_pyxis_filt.loc[idx, 'CalculatedVol_ml'] = convert_mass_to_volume(val, unit, conc_per_ml, conc_unit)
                
                # Convert depending on unit choice
                if order_unit == 'mL':
                    converted_val = convert_mass_to_volume(val, unit, conc_per_ml, conc_unit)
                    df_pyxis_filt.loc[idx, 'CalculatedAmt_RecUnit'] = converted_val
                else:
                    df_pyxis_filt.loc[idx, 'CalculatedAmt_RecUnit'] = val
                    
        vends = group[(group['TransactionType_Clean'] == 'Vend') & (group['TransactionStatus'] == 'Active')].copy()
        wastes = group[(group['TransactionType_Clean'] == 'Waste') & (group['TransactionStatus'] == 'Active')].copy()
        
        # Perform closest-timestamp pairing (greedy bipartite matching, no limit)
        paired_vends = set()
        wastes_sorted = wastes.sort_values('TransactionDateTime_Clean')
        for w_idx, w_row in wastes_sorted.iterrows():
            w_time = w_row['TransactionDateTime_Clean']
            w_val = df_pyxis_filt.loc[w_idx, 'CalculatedAmt_RecUnit']
            
            best_v_idx = None
            min_diff_hours = float('inf')
            for v_idx, v_row in vends.iterrows():
                v_tx_idx = v_row['TxIndex']
                if v_tx_idx in paired_vends:
                    continue
                v_time = v_row['TransactionDateTime_Clean']
                if w_time and v_time:
                    diff_hours = abs((w_time - v_time).total_seconds()) / 3600.0
                    if diff_hours < min_diff_hours:
                        min_diff_hours = diff_hours
                        best_v_idx = v_idx
                        
            if best_v_idx is not None:
                v_row = vends.loc[best_v_idx]
                v_tx_idx = v_row['TxIndex']
                paired_vends.add(v_tx_idx)
                v_val = df_pyxis_filt.loc[best_v_idx, 'CalculatedAmt_RecUnit']
                
                df_pyxis_filt.loc[w_idx, 'PairingStatus'] = 'Matched'
                df_pyxis_filt.loc[w_idx, 'PairedTxTime'] = v_row['TransactionDateTime_Clean']
                df_pyxis_filt.loc[w_idx, 'PairedTxType'] = 'Vend'
                df_pyxis_filt.loc[w_idx, 'PairedTxAmount'] = f"{v_val:.2f} {order_unit}" if v_val is not None else "N/A"
                df_pyxis_filt.loc[w_idx, 'PairingGapHours'] = min_diff_hours
                
                df_pyxis_filt.loc[best_v_idx, 'PairingStatus'] = 'Matched'
                df_pyxis_filt.loc[best_v_idx, 'PairedTxTime'] = w_time
                df_pyxis_filt.loc[best_v_idx, 'PairedTxType'] = 'Waste'
                df_pyxis_filt.loc[best_v_idx, 'PairedTxAmount'] = f"{w_val:.2f} {order_unit}" if w_val is not None else "N/A"
                df_pyxis_filt.loc[best_v_idx, 'PairingGapHours'] = min_diff_hours
                
                unadministered = None
                if v_val is not None and w_val is not None:
                    unadministered = v_val - w_val
                    
                paired_bags.append({
                    'OrderNumber': order_num,
                    'PatientID': w_row['PatientID_Clean'],
                    'PatientName': w_row['PatientName_Clean'],
                    'Medication': med_desc,
                    'ReconciliationUnit': order_unit,
                    'VendTime': v_row['TransactionDateTime_Clean'],
                    'WasteTime': w_time,
                    'TimeGapHours': min_diff_hours,
                    'Dispensed_RecUnit': v_val,
                    'Waste_RecUnit': w_val,
                    'UnadministeredRemainder_RecUnit': unadministered
                })
            else:
                df_pyxis_filt.loc[w_idx, 'PairingStatus'] = 'Unmatched - Excess Waste'
                
        for v_idx, v_row in vends.iterrows():
            v_tx_idx = v_row['TxIndex']
            if v_tx_idx not in paired_vends:
                df_pyxis_filt.loc[v_idx, 'PairingStatus'] = 'Unmatched - Missing Waste'
                
    # Populate unmatched transactions sheet data
    df_unpaired = df_pyxis_filt[df_pyxis_filt['PairingStatus'].str.startswith('Unmatched')].copy()
    df_unpaired['RawAmount'] = df_unpaired.apply(
        lambda r: r['DispenseAmount'] if r['TransactionType_Clean'] == 'Vend' else r['WasteAmount'], axis=1
    )
    df_unpaired_out = df_unpaired[[
        'OrderNumber_Clean', 'PatientID_Clean', 'PatientName_Clean', 'TransactionDateTime_Clean',
        'TransactionType_Clean', 'MedDescription', 'RawAmount', 'CalculatedAmt_RecUnit', 'RecUnit', 'PairingStatus'
    ]].rename(columns={
        'OrderNumber_Clean': 'OrderNumber',
        'PatientID_Clean': 'PatientID',
        'PatientName_Clean': 'PatientName',
        'TransactionDateTime_Clean': 'TransactionDateTime',
        'TransactionType_Clean': 'TransactionType',
        'CalculatedAmt_RecUnit': 'Calculated Amount',
        'RecUnit': 'Unit'
    })
    
    df_paired_bags = pd.DataFrame(paired_bags)
    
    # Only consider matched Order Id rows for all calculations
    matched_order_ids = {r['OrderID'] for r in matching_rows if r['MatchStatus'] == 'Matched'}
    if not df_paired_bags.empty:
        df_paired_bags = df_paired_bags[df_paired_bags['OrderNumber'].astype(str).isin(matched_order_ids)].copy()
        
    if df_paired_bags.empty:
        df_paired_bags = pd.DataFrame(columns=[
            'OrderNumber', 'PatientID', 'PatientName', 'Medication', 'ReconciliationUnit', 'VendTime',
            'WasteTime', 'TimeGapHours', 'Dispensed_RecUnit', 'Waste_RecUnit', 'UnadministeredRemainder_RecUnit'
        ])
        
    reconciliation_summary = []
    administered_summary = []
    paired_order_ids = set()
    if not df_paired_bags.empty:
        paired_order_ids = set(df_paired_bags['OrderNumber'].astype(str).unique())
        
    unpaired_order_ids = set()
    if not df_unpaired.empty:
        unpaired_order_ids = set(df_unpaired['OrderNumber_Clean'].astype(str).unique())
        
    # Reconcile all orders that have any activity in the Pyxis cabinet this week
    all_pyxis_order_ids = sorted(list(set(df_pyxis_filt['OrderNumber_Clean'].unique()) - {''}))

    # -----------------------------------------------------------------------
    # PRE-LOOP: Identify overlapping Order IDs and pair previous-week records
    # upfront so that ALL calculations start with complete, paired history.
    # -----------------------------------------------------------------------
    prev_paired_data = {}   # order_id -> fully-paired prev-week DataFrame

    if not df_pyxis_prev.empty:
        prev_order_ids = set(df_pyxis_prev['OrderNumber_Clean'].dropna().unique()) - {''}
        overlapping_order_ids = set(all_pyxis_order_ids) & prev_order_ids
        print(f"  Overlapping Order IDs (current week AND previous week): {len(overlapping_order_ids)}")

        for ov_order_id in overlapping_order_ids:
            # Get previous-week records for this order
            ov_prev_raw = df_pyxis_prev[df_pyxis_prev['OrderNumber_Clean'] == ov_order_id].copy()
            if ov_prev_raw.empty:
                continue

            # Fetch concentration/unit info resolved during current-week pairing
            ov_conc_data = order_concentrations.get(ov_order_id, {'conc_per_ml': None, 'conc_unit': None})
            ov_conc_per_ml = ov_conc_data.get('conc_per_ml')
            ov_conc_unit  = ov_conc_data.get('conc_unit')
            ov_r_info = order_recon_info.get(ov_order_id, {'unit': 'mL', 'has_bolus': False})
            ov_unit = ov_r_info['unit']

            # Parse and convert amounts using same logic as current-week
            def _prev_parse_row(row, _cpm=ov_conc_per_ml, _cu=ov_conc_unit, _ou=ov_unit):
                tx = row['TransactionType_Clean']
                raw = row['DispenseAmount'] if tx == 'Vend' else row['WasteAmount']
                val, unit = parse_amount_string(raw)
                vol = convert_mass_to_volume(val, unit, _cpm, _cu)
                if _ou == 'mL':
                    calc_amt = vol
                else:
                    calc_amt = val
                return pd.Series([val, vol, calc_amt])

            ov_prev_raw[['CalculatedDose', 'CalculatedVol_ml', 'CalculatedAmt_RecUnit']] = ov_prev_raw.apply(_prev_parse_row, axis=1)
            ov_prev_raw['RecUnit'] = ov_unit
            ov_prev_raw['ParsedConc'] = ov_conc_per_ml
            ov_prev_raw['ParsedConcUnit'] = ov_conc_unit
            ov_prev_raw['ParsedTotalVol'] = None
            ov_prev_raw['Audit Week'] = 'Previous Week'
            ov_prev_raw['TxIndex'] = range(len(ov_prev_raw))

            # Initialise pairing columns
            ov_prev_raw['PairingStatus'] = 'Unmatched'
            ov_prev_raw.loc[ov_prev_raw['TransactionStatus'].isin(['Stock Return', 'Cancelled']), 'PairingStatus'] = 'Excluded'
            ov_prev_raw['PairedTxTime'] = None
            ov_prev_raw['PairedTxType'] = None
            ov_prev_raw['PairedTxAmount'] = None
            ov_prev_raw['PairingGapHours'] = None

            # Greedy Vend→Waste pairing on previous-week records
            ov_vends  = ov_prev_raw[(ov_prev_raw['TransactionType_Clean'] == 'Vend')  & (ov_prev_raw['TransactionStatus'] == 'Active')].copy()
            ov_wastes = ov_prev_raw[(ov_prev_raw['TransactionType_Clean'] == 'Waste') & (ov_prev_raw['TransactionStatus'] == 'Active')].copy()

            ov_paired_vends = set()
            for w_idx, w_row in ov_wastes.sort_values('TransactionDateTime_Clean').iterrows():
                w_time = w_row['TransactionDateTime_Clean']
                w_val  = w_row['CalculatedAmt_RecUnit']
                best_v_idx = None
                min_gap    = float('inf')
                for v_idx, v_row in ov_vends.iterrows():
                    v_tx = v_row['TxIndex']
                    if v_tx in ov_paired_vends:
                        continue
                    v_time = v_row['TransactionDateTime_Clean']
                    if w_time and v_time:
                        gap = abs((w_time - v_time).total_seconds()) / 3600.0
                        if gap < min_gap:
                            min_gap = gap
                            best_v_idx = v_idx
                if best_v_idx is not None:
                    v_row = ov_vends.loc[best_v_idx]
                    ov_paired_vends.add(v_row['TxIndex'])
                    v_val = v_row['CalculatedAmt_RecUnit']
                    # Mark waste row
                    ov_prev_raw.loc[w_idx, 'PairingStatus']    = 'Matched'
                    ov_prev_raw.loc[w_idx, 'PairedTxTime']     = v_row['TransactionDateTime_Clean']
                    ov_prev_raw.loc[w_idx, 'PairedTxType']     = 'Vend'
                    ov_prev_raw.loc[w_idx, 'PairedTxAmount']   = f"{v_val:.2f} {ov_unit}" if v_val is not None else 'N/A'
                    ov_prev_raw.loc[w_idx, 'PairingGapHours']  = min_gap
                    # Mark vend row
                    ov_prev_raw.loc[best_v_idx, 'PairingStatus']    = 'Matched'
                    ov_prev_raw.loc[best_v_idx, 'PairedTxTime']     = w_time
                    ov_prev_raw.loc[best_v_idx, 'PairedTxType']     = 'Waste'
                    ov_prev_raw.loc[best_v_idx, 'PairedTxAmount']   = f"{w_val:.2f} {ov_unit}" if w_val is not None else 'N/A'
                    ov_prev_raw.loc[best_v_idx, 'PairingGapHours']  = min_gap
                    # Add to paired bags
                    unadmin = (v_val - w_val) if (v_val is not None and w_val is not None) else None
                    paired_bags.append({
                        'OrderNumber': ov_order_id,
                        'PatientID': w_row.get('PatientID_Clean', ''),
                        'PatientName': w_row.get('PatientName_Clean', ''),
                        'Medication': w_row.get('MedDescription', ''),
                        'ReconciliationUnit': ov_unit,
                        'VendTime': v_row['TransactionDateTime_Clean'],
                        'WasteTime': w_time,
                        'TimeGapHours': min_gap,
                        'Dispensed_RecUnit': v_val,
                        'Waste_RecUnit': w_val,
                        'UnadministeredRemainder_RecUnit': unadmin
                    })
                else:
                    ov_prev_raw.loc[w_idx, 'PairingStatus'] = 'Unmatched - Excess Waste'

            for v_idx, v_row in ov_vends.iterrows():
                if v_row['TxIndex'] not in ov_paired_vends:
                    ov_prev_raw.loc[v_idx, 'PairingStatus'] = 'Unmatched - Missing Waste'

            prev_paired_data[ov_order_id] = ov_prev_raw

    # -----------------------------------------------------------------------
    # END PRE-LOOP
    # -----------------------------------------------------------------------

    for order_id in all_pyxis_order_ids:
        # Retrieve Order info
        r_info = order_recon_info.get(order_id, {'unit': 'mL', 'has_bolus': False})
        order_unit = r_info['unit']
        has_bolus = r_info['has_bolus']
        
        p_prev_sub_filt = pd.DataFrame()
        p_sub = df_pyxis_filt[df_pyxis_filt['OrderNumber_Clean'] == order_id]
        
        e_this = df_epic_this[df_epic_this['OrderID_Clean'] == order_id] if not df_epic_this.empty else pd.DataFrame()
        e_prev = df_epic_prev[df_epic_prev['OrderID_Clean'] == order_id] if not df_epic_prev.empty else pd.DataFrame()
        f_this = df_flow_this[df_flow_this['OrderID_Clean'] == order_id] if not df_flow_this.empty else pd.DataFrame()
        f_prev = df_flow_prev[df_flow_prev['OrderID_Clean'] == order_id] if not df_flow_prev.empty else pd.DataFrame()
        
        # Combine all previous and current week records for calculations (no window filtering)
        f_prev_filt = f_prev
        e_sub = pd.concat([e_prev, e_this], ignore_index=True)
        f_sub = pd.concat([f_prev, f_this], ignore_index=True)
        
        # Look up pre-paired previous-week records built upfront before the loop
        p_prev_sub_filt = prev_paired_data.get(order_id, pd.DataFrame())
        if not p_prev_sub_filt.empty:
            prev_pyxis_rows.append(p_prev_sub_filt)
        
        pyxis_patient = p_sub.iloc[0]['PatientName_Clean'] if not p_sub.empty else "N/A"
        pyxis_med = p_sub.iloc[0]['MedDescription'] if not p_sub.empty else "N/A"
        
        conc_data = order_concentrations.get(order_id, {'conc_per_ml': None, 'conc_unit': None})
        conc_per_ml = conc_data['conc_per_ml']
        
        # Mapped Flowshet drug name
        flowsheet_drug_row = None
        if pyxis_med != "N/A":
            desc_lower = pyxis_med.lower()
            if 'ketamine' in desc_lower:
                flowsheet_drug_row = 'R KETAMINE VOLUME'
            elif 'midazolam' in desc_lower or 'versed' in desc_lower:
                flowsheet_drug_row = 'R MIDAZOLAM VOLUME'
            elif 'fentanyl' in desc_lower:
                flowsheet_drug_row = 'R FENTANYL VOLUME'
            elif 'morphine' in desc_lower:
                flowsheet_drug_row = 'R MORPHINE VOLUME'
            elif 'hydromorphone' in desc_lower or 'dilaudid' in desc_lower:
                flowsheet_drug_row = 'R HYDROMORPHONE VOLUME'
                
        # Check flowsheet row name matching (AI synonym match rule)
        f_matching_rows = pd.DataFrame()
        if not f_sub.empty and flowsheet_drug_row:
            f_matching_rows = f_sub[f_sub['Flowsheet Row Name'] == flowsheet_drug_row].copy()
            f_unmatched_rows = f_sub[f_sub['Flowsheet Row Name'] != flowsheet_drug_row].copy()
            
            for _, r in f_unmatched_rows.iterrows():
                exceptions.append({
                    'OrderID': order_id,
                    'PatientName': pyxis_patient,
                    'Source': 'Epic Flowsheets',
                    'ExceptionType': 'AI Medication Match Exception',
                    'Reason': f"Flowsheet Row Name '{r['Flowsheet Row Name']}' on Order {order_id} does not match Pyxis medication '{pyxis_med}'. Confidence=0.0",
                    'ActionNeeded': 'Verify if flowsheet records correct drug; confirm generic substitution rules.'
                })
        elif not f_sub.empty and not flowsheet_drug_row:
            for _, r in f_sub.iterrows():
                exceptions.append({
                    'OrderID': order_id,
                    'PatientName': pyxis_patient,
                    'Source': 'Epic Flowsheets',
                    'ExceptionType': 'AI Medication Match Exception',
                    'Reason': f"Unrecognized Pyxis medication '{pyxis_med}'; flowsheet row '{r['Flowsheet Row Name']}' match confidence=0.0",
                    'ActionNeeded': 'Confirm pharmacy drug mapping settings.'
                })
                
        # Calculate totals from current and previous week's Pyxis transactions (if considered)
        p_sub_all = pd.concat([p_prev_sub_filt, p_sub], ignore_index=True) if not p_prev_sub_filt.empty else p_sub
        curr_disp_dose = p_sub_all[p_sub_all['TransactionType_Clean'] == 'Vend']['CalculatedDose'].dropna().sum()
        curr_disp_vol = p_sub_all[p_sub_all['TransactionType_Clean'] == 'Vend']['CalculatedVol_ml'].dropna().sum()
        curr_waste_dose = p_sub_all[p_sub_all['TransactionType_Clean'] == 'Waste']['CalculatedDose'].dropna().sum()
        curr_waste_vol = p_sub_all[p_sub_all['TransactionType_Clean'] == 'Waste']['CalculatedVol_ml'].dropna().sum()
        
        # Flowsheet Administered
        curr_admin_vol = 0.0
        curr_admin_dose = 0.0
        
        if not f_matching_rows.empty:
            curr_admin_vol = f_matching_rows['FlowsheetValue_Num'].sum()
            
            for _, f_row in f_matching_rows.iterrows():
                conv_dose = ""
                if conc_per_ml is not None:
                    conv_val = f_row['FlowsheetValue_Num'] * conc_per_ml
                    conv_dose = f"{conv_val:.2f} {order_unit}"
                    if order_unit != 'mL':
                        curr_admin_dose += conv_val
                else:
                    conv_dose = "Missing Concentration"
                    
                flowsheet_log_rows.append({
                    'OrderID': order_id,
                    'CSN': f_row['CSN'],
                    'MRN': f_row['MRN_Clean'],
                    'PatientName': f_row['Patient Name'],
                    'FlowsheetRowName': f_row['Flowsheet Row Name'],
                    'RecordedTime': f_row['RecordedTime_Clean'],
                    'InfusionVolume_mL': f_row['FlowsheetValue_Num'],
                    'ConvertedDose': conv_dose if order_unit != 'mL' else "",
                    'Unit': order_unit,
                    'Status': 'Confidently Matched (Confidence=1.0)',
                    'Audit Week': f_row.get('Audit Week', 'Current Week')
                })
                
        # Check if concentration is missing for bolus conversion
        if has_bolus and (conc_per_ml is None or conc_per_ml == 0):
            exceptions.append({
                'OrderID': order_id,
                'PatientName': pyxis_patient,
                'Source': 'Reconciliation Logic',
                'ExceptionType': 'Missing Concentration Data',
                'Reason': f"Bolus order requires concentration conversion, but Pyxis description '{pyxis_med}' concentration is missing.",
                'ActionNeeded': 'Manually verify drug concentration from pharmacy records and reconcile dose.'
            })
            
        # Select target reconciliation unit values
        if order_unit == 'mL':
            acc_dispensed = curr_disp_vol
            acc_administered = curr_admin_vol
            acc_waste = curr_waste_vol
        else:
            acc_dispensed = curr_disp_dose
            acc_administered = curr_admin_dose
            acc_waste = curr_waste_dose
            
        expected_waste = clean_val_for_json(acc_dispensed - acc_administered)
        unadmin_remainder = clean_val_for_json(acc_dispensed - acc_waste)
        variance_waste = clean_val_for_json(expected_waste - acc_waste)
        variance_given = clean_val_for_json(unadmin_remainder - acc_administered)
        
        # Resolve MRN
        mrn = "N/A"
        if not p_sub.empty:
            p_mrn = p_sub['PatientID_Clean'].dropna()
            if not p_mrn.empty:
                mrn = p_mrn.iloc[0]
        if mrn == "N/A" and not e_sub.empty:
            e_mrn = e_sub['MRN_Clean'].dropna()
            if not e_mrn.empty:
                mrn = e_mrn.iloc[0]
        if mrn == "N/A" and not f_sub.empty:
            f_mrn = f_sub['MRN_Clean'].dropna()
            if not f_mrn.empty:
                mrn = f_mrn.iloc[0]
                
        total_bags = len(p_sub[(p_sub['TransactionType_Clean'] == 'Vend') & (p_sub['TransactionStatus'] == 'Active')])
        infusion_type = "Bolus" if has_bolus else "Continuous Infusion"
        variance_pct = abs(variance_waste) / acc_dispensed if acc_dispensed > 0.0 else 0.0
        
        # Check Stopped action in current week administrations
        has_stopped = False
        if not e_sub.empty:
            has_stopped = any('stop' in str(a).lower() for a in e_sub['AdministrationAction'])
            
        # Combine timestamps for this week's duration
        t_list = []
        for t in p_sub['TransactionDateTime_Clean'].dropna().tolist():
            t_list.append(t)
        for t in e_sub['AdministrationInstant_Clean'].dropna().tolist():
            t_list.append(t)
        if not f_matching_rows.empty:
            for t in f_matching_rows['RecordedTime_Clean'].dropna().tolist():
                t_list.append(t)
                
        elapsed_hours_total = 0.0
        if t_list:
            elapsed_hours_total = (max(t_list) - min(t_list)).total_seconds() / 3600.0
            
        order_status = 'Active - Continuing Next Week'
        if has_stopped:
            order_status = 'Completed'
        elif elapsed_hours_total > 96.0:
            order_status = 'Completed - Max Hang Time'
            
        # Determine status messages
        if order_status == 'Active - Continuing Next Week':
            rec_status = "Active - Pending Cross-Week"
            var_alert = "No"
            audit_note = "Infusion order remains active in clinical flowsheet. Reconcile on completion."
        else:
            if abs(variance_waste) <= 0.1:
                rec_status = "Reconciled"
                var_alert = "No"
                audit_note = "Reconciliation complete. Documented waste matches expected volumes."
            else:
                rec_status = "Discrepancy Alert"
                var_alert = "Yes"
                audit_note = f"Material variance of {variance_waste:.2f} {order_unit} detected upon order completion."
                
                # Log Variance exceptions
                exceptions.append({
                    'OrderID': order_id,
                    'PatientName': pyxis_patient,
                    'Source': 'Reconciliation Logic',
                    'ExceptionType': 'Material Reconciliation Variance',
                    'Reason': f"Completed order has a variance of {variance_waste:.2f} {order_unit} (Expected Waste={expected_waste:.2f}, Documented={acc_waste:.2f}).",
                    'ActionNeeded': 'Audit pump flowsheet charts, cabinet dispenses, and waste entries to locate discrepancy.'
                })
                
        # Check for missing waste documentation (expected waste > 0.1 but waste is 0)
        if expected_waste > 0.1 and acc_waste <= 0.1:
            exceptions.append({
                'OrderID': order_id,
                'PatientName': pyxis_patient,
                'Source': 'Reconciliation Logic',
                'ExceptionType': 'Missing Waste Documentation',
                'Reason': f"Expected Waste is {expected_waste:.2f} {order_unit} but Documented Waste is {acc_waste:.2f} {order_unit}.",
                'ActionNeeded': 'Verify if unused medication was wasted or returned; audit clinician waste documentation.'
            })
            
        # Check for excess waste documentation (documented waste exceeds expected waste)
        if variance_waste < -0.1:
            exceptions.append({
                'OrderID': order_id,
                'PatientName': pyxis_patient,
                'Source': 'Reconciliation Logic',
                'ExceptionType': 'Excess Waste Documentation',
                'Reason': f"Documented Waste ({acc_waste:.2f} {order_unit}) exceeds Expected Waste ({expected_waste:.2f} {order_unit}).",
                'ActionNeeded': 'Check if excess wasting transactions were logged in Pyxis or if flowsheet admins were under-charted.'
            })
            
        # Check for missing administrations (dispensed but 0 administered)
        if acc_dispensed > 0.1 and acc_administered <= 0.0:
            exceptions.append({
                'OrderID': order_id,
                'PatientName': pyxis_patient,
                'Source': 'Epic Flowsheets',
                'ExceptionType': 'Missing Administration Activity',
                'Reason': f"Order has cabinet dispenses ({acc_dispensed:.2f} {order_unit}) but no administration recorded in flowsheet.",
                'ActionNeeded': 'Check if infusion was started or if flowsheet volumes were charted under a different order.'
            })
            
        # Order Hang Time Exceeded
        if order_status == 'active' and elapsed_hours_total > 96.0:
            exceptions.append({
                'OrderID': order_id,
                'PatientName': pyxis_patient,
                'Source': 'Reconciliation Logic',
                'ExceptionType': 'Order Hang Time Exceeded',
                'Reason': f"Order remains active after {elapsed_hours_total:.1f} hours without a Stopped action.",
                'ActionNeeded': 'Confirm with clinical staff if infusion was discontinued; manually close order in audit.'
            })
            
        has_unmatched = "Yes" if order_id in unpaired_order_ids else "No"
        admin_status = "Missing Admin Charting" if acc_administered <= 0.0 else "Charted"
        
        # Get unique list of charted actions
        admin_actions_str = "N/A"
        if not e_sub.empty:
            seen_act = set()
            unique_actions = [a for a in e_sub['AdministrationAction'].astype(str).str.strip().tolist() if not (a in seen_act or seen_act.add(a))]
            admin_actions_str = ", ".join(unique_actions)
            
        # Determine Cross-Week Status
        is_prev_considered = not p_prev_sub_filt.empty
        is_continues_next = (order_status == 'Active - Continuing Next Week')
            
        if is_prev_considered and is_continues_next:
            cross_week_status = 'Considered Previous Week & Continues Next Week'
        elif is_prev_considered:
            cross_week_status = 'Considered Previous Week Data'
        elif is_continues_next:
            cross_week_status = 'Continues Next Week'
        else:
            cross_week_status = 'Current Week Only'
            
        reconciliation_summary.append({
            'OrderID': order_id,
            'PatientName': pyxis_patient,
            'Medication': pyxis_med,
            'ReconciliationUnit': order_unit,
            'TotalDispensed': acc_dispensed,
            'TotalWaste': acc_waste,
            'CumulativeUnadministeredRemainder': unadmin_remainder,
            'TotalAdministered': acc_administered,
            'VarianceInGiven': variance_given,
            'OrderStatus': order_status,
            'HasUnmatchedCabinetPairs': has_unmatched,
            'CumulativeAdministeredStatus': admin_status,
            'AdministrationAction': admin_actions_str,
            'Cross-Week Status': cross_week_status
        })
        
        administered_summary.append({
            'OrderID': order_id,
            'PatientName': pyxis_patient,
            'Medication': pyxis_med,
            'ReconciliationUnit': order_unit,
            'FlowsheetEntriesCount': len(f_matching_rows),
            'TotalVolume_mL': curr_admin_vol,
            'ConcentrationPerML': conc_per_ml if conc_per_ml is not None else 0.0,
            'ConvertedDose': curr_admin_dose,
            'FinalCumulativeAdministered': acc_administered
        })
        
    # Log exceptions for Pyxis transactions with missing Order ID
    p_missing_order = df_pyxis_filt[
        (df_pyxis_filt['OrderNumber_Clean'] == '') & 
        (df_pyxis_filt['TransactionStatus'] == 'Active')
    ]
    for _, row in p_missing_order.iterrows():
        tx_type = row['TransactionType_Clean']
        raw_amt = row['DispenseAmount'] if tx_type == 'Vend' else row['WasteAmount']
        exceptions.append({
            'OrderID': 'MISSING',
            'PatientName': row['PatientName_Clean'],
            'Source': 'Pyxis Cabinets',
            'ExceptionType': 'Missing Order ID in Pyxis',
            'Reason': f"Cabinet {tx_type} transaction of {raw_amt} has no Order ID documented in Pyxis.",
            'ActionNeeded': 'Audit Pyxis cabinet access log and patient charts to link transaction to a valid clinical order.'
        })
        
    df_matching_overview = pd.DataFrame(matching_rows)
    df_reconciliation_summary = pd.DataFrame(reconciliation_summary)
    df_administered_summary = pd.DataFrame(administered_summary)
    df_flowsheet_log = pd.DataFrame(flowsheet_log_rows)
    if df_flowsheet_log.empty:
        df_flowsheet_log = pd.DataFrame(columns=[
            'OrderID', 'CSN', 'MRN', 'PatientName', 'FlowsheetRowName', 'RecordedTime', 'InfusionVolume_mL',
            'ConvertedDose', 'Unit', 'Status', 'Audit Week'
        ])
        
    df_exceptions = pd.DataFrame(exceptions)
    if df_exceptions.empty:
        df_exceptions = pd.DataFrame(columns=['OrderID', 'PatientName', 'Source', 'ExceptionType', 'Reason', 'ActionNeeded'])
        

        
    # Prep Dashboard values
    total_pyxis = len(df_pyxis_filt)
    total_epic = len(df_epic)
    paired_bags_count = len(df_paired_bags)
    
    # Calculate Dashboard stats
    orders_with_exceptions = {ex.get('OrderID') for ex in exceptions if ex.get('OrderID') and ex.get('OrderID') != 'MISSING'}
    completed_orders = len(df_reconciliation_summary[df_reconciliation_summary['OrderStatus'].str.startswith('Completed')])
    active_orders = len(df_reconciliation_summary[df_reconciliation_summary['OrderStatus'] == 'Active'])
    reconciled_count = len(df_reconciliation_summary[
        (df_reconciliation_summary['OrderStatus'].str.startswith('Completed')) & 
        (df_reconciliation_summary['VarianceInGiven'].abs() <= 0.1)
    ])
    discrepancy_count = len(df_reconciliation_summary[
        (df_reconciliation_summary['OrderStatus'].str.startswith('Completed')) & 
        (df_reconciliation_summary['VarianceInGiven'].abs() > 0.1)
    ])
    total_exceptions_count = len(df_exceptions)
    
    print("\nPairing & Reconciliation Statistics:")
    print(f"  Total Pyxis Transactions   : {total_pyxis}")
    print(f"  Total Epic Admins          : {total_epic}")
    print(f"  Successfully Paired Bags   : {paired_bags_count}")
    print(f"  Active (Pending) Orders    : {active_orders}")
    print(f"  Completed Reconciled Orders: {reconciled_count}")
    print(f"  Completed Discrepant Orders: {discrepancy_count}")
    print(f"  Total Audit Exceptions     : {total_exceptions_count}")
    
    # 7. Generate Expanded Excel Report
    print(f"\nGenerating Excel Report at: {OUTPUT_PATH}...")
    wb = openpyxl.Workbook()
    
    # Summary Dashboard removed
        
    # Formatting Helper
    def format_sheet(ws, headers, df_data, highlight_rules=None, header_comments=None):
        ws.views.sheetView[0].showGridLines = True
        
        # Write headers to row 1
        for col_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=h)
            cell.fill = NAVY_HEADER_FILL
            cell.font = NAVY_HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = THIN_BORDER
            
            if header_comments and h in header_comments:
                from openpyxl.comments import Comment
                comment = Comment(text=header_comments[h], author="System")
                comment.width = 250
                comment.height = 60
                cell.comment = comment
        ws.row_dimensions[1].height = 28
        
        # Write data starting from row 2
        for r_idx, row_vals in enumerate(df_data.values, start=2):
            ws.row_dimensions[r_idx].height = 20
            fill = ZEBRA_FILL if r_idx % 2 == 0 else WHITE_FILL
            for c_idx, val in enumerate(row_vals, start=1):
                cell = ws.cell(row=r_idx, column=c_idx)
                
                # Check if it is an ID column based on header name
                h_name = str(headers[c_idx - 1]).lower().strip()
                is_id_col = False
                if any(x in h_name for x in ["order", "patient id", "mrn", "patientid", "csn"]):
                    is_id_col = True
                    
                if is_id_col and val is not None and not pd.isna(val):
                    # Format as clean text string
                    if isinstance(val, float):
                        val_str = str(int(val)) if val.is_integer() else f"{val:.0f}"
                    else:
                        val_str = str(val).strip()
                        if val_str.endswith('.0'):
                            val_str = val_str[:-2]
                    cell.value = val_str
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.number_format = '@'
                    cell.font = DATA_FONT
                    cell.fill = fill
                    cell.border = THIN_BORDER
                else:
                    cell.value = None if (val is None or pd.isna(val)) else val
                    cell.font = DATA_FONT
                    cell.fill = fill
                    cell.border = THIN_BORDER
                    
                    if val is not None and not pd.isna(val):
                        if isinstance(val, (int, float)):
                            if "percentage" in h_name or "pct" in h_name or "%" in h_name:
                                cell.number_format = '0.0%'
                            elif "duration" in h_name or "hours" in h_name:
                                cell.number_format = '#,##0.0'
                            else:
                                cell.number_format = '#,##0.00' if isinstance(val, float) else '#,##0'
                            cell.alignment = Alignment(horizontal="right", vertical="center")
                        elif hasattr(val, 'strftime'):
                            cell.value = val.strftime('%Y-%m-%d %H:%M')
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                        else:
                            cell.alignment = Alignment(horizontal="left", vertical="center")
                            v_str = str(val).strip()
                            if v_str.isdigit() and len(v_str) > 5:
                                cell.alignment = Alignment(horizontal="center", vertical="center")
                                cell.number_format = '@'
                        
            if highlight_rules:
                highlight_rules(ws, r_idx, row_vals)
                
        # Sizing
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col[1:]:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    def format_stacked_tables(ws, title_flow, flow_headers, df_flow, title_admin, admin_headers, df_admin):
        ws.views.sheetView[0].showGridLines = True
        
        # Write Title 1
        cell_t1 = ws.cell(row=1, column=1, value=title_flow)
        cell_t1.font = Font(name="Segoe UI", size=12, bold=True, color="1B365D")
        ws.row_dimensions[1].height = 24
        
        # Write Flowsheet Headers
        for col_idx, h in enumerate(flow_headers, start=1):
            cell = ws.cell(row=2, column=col_idx, value=h)
            cell.fill = NAVY_HEADER_FILL
            cell.font = NAVY_HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = THIN_BORDER
        ws.row_dimensions[2].height = 28
        
        # Write Flowsheet Data
        curr_row = 3
        if not df_flow.empty:
            for r_idx, row_vals in enumerate(df_flow.values):
                ws.row_dimensions[curr_row].height = 20
                fill = ZEBRA_FILL if curr_row % 2 == 0 else WHITE_FILL
                for c_idx, val in enumerate(row_vals, start=1):
                    cell = ws.cell(row=curr_row, column=c_idx)
                    h_name = str(flow_headers[c_idx - 1]).lower().strip()
                    is_id = any(x in h_name for x in ["order", "patient id", "mrn", "csn"])
                    
                    if is_id and val is not None and not pd.isna(val):
                        val_str = str(val).strip()
                        if val_str.endswith('.0'): val_str = val_str[:-2]
                        cell.value = val_str
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        cell.number_format = '@'
                    else:
                        cell.value = None if (val is None or pd.isna(val)) else val
                        if isinstance(val, (int, float)) and not pd.isna(val):
                            cell.number_format = '0.00'
                    cell.font = DATA_FONT
                    cell.fill = fill
                    cell.border = THIN_BORDER
                curr_row += 1
        else:
            ws.cell(row=curr_row, column=1, value="No flowsheet records considered").font = Font(name="Segoe UI", size=10, italic=True)
            ws.row_dimensions[curr_row].height = 20
            curr_row += 1
            
        curr_row += 2 # Spacer
        
        # Write Title 2
        cell_t2 = ws.cell(row=curr_row, column=1, value=title_admin)
        cell_t2.font = Font(name="Segoe UI", size=12, bold=True, color="1B365D")
        ws.row_dimensions[curr_row].height = 24
        curr_row += 1
        
        # Write Admin Headers
        for col_idx, h in enumerate(admin_headers, start=1):
            cell = ws.cell(row=curr_row, column=col_idx, value=h)
            cell.fill = NAVY_HEADER_FILL
            cell.font = NAVY_HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = THIN_BORDER
        ws.row_dimensions[curr_row].height = 28
        curr_row += 1
        
        # Write Admin Data
        if not df_admin.empty:
            for r_idx, row_vals in enumerate(df_admin.values):
                ws.row_dimensions[curr_row].height = 20
                fill = ZEBRA_FILL if curr_row % 2 == 0 else WHITE_FILL
                for c_idx, val in enumerate(row_vals, start=1):
                    cell = ws.cell(row=curr_row, column=c_idx)
                    h_name = str(admin_headers[c_idx - 1]).lower().strip()
                    is_id = any(x in h_name for x in ["order", "patient id", "mrn", "csn"])
                    
                    if is_id and val is not None and not pd.isna(val):
                        val_str = str(val).strip()
                        if val_str.endswith('.0'): val_str = val_str[:-2]
                        cell.value = val_str
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        cell.number_format = '@'
                    else:
                        cell.value = None if (val is None or pd.isna(val)) else val
                        if isinstance(val, (int, float)) and not pd.isna(val):
                            cell.number_format = '0.00'
                    cell.font = DATA_FONT
                    cell.fill = fill
                    cell.border = THIN_BORDER
                    
                    # Highlights for administration action stop/restart
                    if h_name == "administrative action" and val is not None:
                        action_val = str(val).lower()
                        if 'stop' in action_val:
                            cell.fill = RED_FILL
                            cell.font = RED_FONT
                        elif 'restart' in action_val:
                            cell.fill = YELLOW_FILL
                            cell.font = YELLOW_FONT
                curr_row += 1
        else:
            ws.cell(row=curr_row, column=1, value="No administration records considered").font = Font(name="Segoe UI", size=10, italic=True)
            ws.row_dimensions[curr_row].height = 20
            curr_row += 1
            
        # Auto-adjust column widths
        max_len = {}
        # Scan headers
        for col_idx, h in enumerate(flow_headers, start=1):
            max_len[col_idx] = max(max_len.get(col_idx, 0), len(str(h)))
        for col_idx, h in enumerate(admin_headers, start=1):
            max_len[col_idx] = max(max_len.get(col_idx, 0), len(str(h)))
        # Scan data
        for r in range(1, curr_row):
            for c in range(1, max(len(flow_headers), len(admin_headers)) + 1):
                val = ws.cell(row=r, column=c).value
                if val is not None:
                    max_len[c] = max(max_len.get(c, 0), len(str(val)))
        for c, w in max_len.items():
            ws.column_dimensions[get_column_letter(c)].width = max(w + 3, 12)

    ws_recon = wb.active
    ws_recon.title = "Order Reconciliation Summary"
    recon_headers = [
        "Order ID", "Patient Name", "Medication Name", "Reconciliation Unit",
        "Cumulative Dispensed", "Cumulative Initial Waste", "Available to be Administered",
        "Cumulative Administered", "Additional in Waste", "Order Status",
        "Has Unmatched Pyxis Transactions", "Cumulative Administered Status", "Administration Action",
        "Cross-Week Status"
    ]
    def rule_recon(ws, r_idx, row_vals):
        status_val = str(row_vals[9]).lower()
        has_unmatched = str(row_vals[10]).strip()
        admin_status = str(row_vals[11]).strip()
        
        # Color Status (Column 10)
        # Check if additional in waste is very small
        additional_waste_val = 0.0
        try:
            additional_waste_val = float(row_vals[8])
        except:
            pass
        has_discrepancy = abs(additional_waste_val) > 0.1
        
        if 'max hang time' in status_val:
            ws.cell(row=r_idx, column=10).fill = ORANGE_FILL
            ws.cell(row=r_idx, column=10).font = ORANGE_FONT
        elif status_val.startswith('completed'):
            ws.cell(row=r_idx, column=10).fill = GREEN_FILL
            ws.cell(row=r_idx, column=10).font = GREEN_FONT
        elif 'continuing next week' in status_val:
            ws.cell(row=r_idx, column=10).fill = PatternFill(start_color="D6E4FF", end_color="D6E4FF", fill_type="solid")
            ws.cell(row=r_idx, column=10).font = Font(name="Segoe UI", size=10, color="002D62", bold=True)
        else:
            ws.cell(row=r_idx, column=10).fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
            ws.cell(row=r_idx, column=10).font = Font(name="Segoe UI", size=10, color="856404", bold=True)
            
        # Color Unmatched Pairs column (Column 11)
        if has_unmatched == 'Yes':
            ws.cell(row=r_idx, column=11).fill = YELLOW_FILL
            ws.cell(row=r_idx, column=11).font = YELLOW_FONT
        else:
            ws.cell(row=r_idx, column=11).fill = GREEN_FILL
            ws.cell(row=r_idx, column=11).font = GREEN_FONT
            
        # Color Cumulative Administered Status (Column 12)
        if admin_status == 'Missing Admin Charting':
            ws.cell(row=r_idx, column=12).fill = RED_FILL
            ws.cell(row=r_idx, column=12).font = RED_FONT
        else:
            ws.cell(row=r_idx, column=12).fill = GREEN_FILL
            ws.cell(row=r_idx, column=12).font = GREEN_FONT
            
        # Color Cross-Week Status (Column 14)
        cross_week_val = str(row_vals[13]).strip()
        if 'Considered' in cross_week_val or 'Continues' in cross_week_val:
            ws.cell(row=r_idx, column=14).fill = PatternFill(start_color="E6F4EA", end_color="E6F4EA", fill_type="solid")
            ws.cell(row=r_idx, column=14).font = Font(name="Segoe UI", size=10, color="137333", bold=True)
            
    recon_header_comments = {
        "Available to be Administered": "Available to be Administered = Cumulative Dispensed (E) - Cumulative Initial Waste (F)",
        "Additional in Waste": "Additional in Waste = Available to be Administered (G) - Cumulative Administered (H)"
    }
    format_sheet(ws_recon, recon_headers, df_reconciliation_summary, rule_recon, header_comments=recon_header_comments)




    # Sheet 6: Epic Administrations Log
    ws_admin = wb.create_sheet(title="Epic Administrations Log")
    admin_headers = [
        "Order ID", "MRN", "Patient Name", "Medication", "Administrative Action", "Administration Instant",
        "Dose", "Dose Unit", "Patient Weight (kg)", "Audit Week"
    ]
    df_epic_out = df_epic[[
        'OrderID_Clean', 'MRN_Clean', 'PatientName_Clean', 'Medication', 'AdministrationAction',
        'AdministrationInstant_Clean', 'Dose', 'DoseUnit', 'WeightAtRelease_X'
    ]].copy()
    df_epic_out['Audit Week'] = 'Current Week'
    
    # Convert collected flowsheet and administrations rows to DataFrames
    df_prev_collected_admin = pd.concat(prev_week_admins_collected, ignore_index=True) if prev_week_admins_collected else pd.DataFrame()
    if not df_prev_collected_admin.empty:
        df_prev_admin_out = df_prev_collected_admin[[
            'OrderID_Clean', 'MRN_Clean', 'PatientName_Clean', 'Medication', 'AdministrationAction',
            'AdministrationInstant_Clean', 'Dose', 'DoseUnit', 'WeightAtRelease_X'
        ]].copy()
        df_prev_admin_out['Audit Week'] = 'Previous Week'
        df_admin_combined = pd.concat([df_epic_out, df_prev_admin_out], ignore_index=True)
    else:
        df_admin_combined = df_epic_out.copy()
        
    df_admin_combined = df_admin_combined.rename(columns={
        'OrderID_Clean': 'Order ID',
        'MRN_Clean': 'MRN',
        'PatientName_Clean': 'Patient Name',
        'AdministrationAction': 'Administrative Action',
        'AdministrationInstant_Clean': 'Administration Instant',
        'WeightAtRelease_X': 'Patient Weight (kg)'
    })
    
    def rule_admin(ws, r_idx, row_vals):
        action_val = str(row_vals[4]).lower()
        if 'stop' in action_val:
            ws.cell(row=r_idx, column=5).fill = RED_FILL
            ws.cell(row=r_idx, column=5).font = RED_FONT
        elif 'restart' in action_val:
            ws.cell(row=r_idx, column=5).fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
            ws.cell(row=r_idx, column=5).font = Font(name="Segoe UI", size=10, color="856404", bold=True)
            
        # Highlight Previous Week rows
        week_val = str(row_vals[9]).strip()
        if week_val == 'Previous Week':
            ws.cell(row=r_idx, column=10).fill = PatternFill(start_color="E6F4EA", end_color="E6F4EA", fill_type="solid")
            ws.cell(row=r_idx, column=10).font = Font(name="Segoe UI", size=10, color="137333", bold=True)
            
    format_sheet(ws_admin, admin_headers, df_admin_combined, rule_admin)

    # Sheet 7: Epic Flowsheet Log
    ws_flow = wb.create_sheet(title="Epic Flowsheet Log")
    flow_headers = [
        "Order ID", "CSN", "MRN", "Patient Name", "Flowsheet Row Name", "Recorded Time",
        "Infusion Volume (mL)", "Converted Dose (Rec Unit)", "Unit", "Mapped Drug Status", "Audit Week"
    ]
    
    def rule_flow(ws, r_idx, row_vals):
        # Highlight Previous Week rows
        week_val = str(row_vals[10]).strip()
        if week_val == 'Previous Week':
            ws.cell(row=r_idx, column=11).fill = PatternFill(start_color="E6F4EA", end_color="E6F4EA", fill_type="solid")
            ws.cell(row=r_idx, column=11).font = Font(name="Segoe UI", size=10, color="137333", bold=True)
            
    format_sheet(ws_flow, flow_headers, df_flowsheet_log, rule_flow)

    # Sheet 9: CS continuous infusions
    ws_pyxis_sheet = wb.create_sheet(title="CS continuous infusions")
    raw_headers = [
        "Order Number", "Patient ID", "Patient Name", "Transaction DateTime", "Transaction Type",
        "Transaction Status", "MedDescription", "DispenseAmount (Raw)", "WasteAmount (Raw)", "Parsed Concentration", "Parsed Unit",
        "Parsed Total Vol", "Calculated Dose (mg/mcg)", "Converted Vol (mL)", "Pairing Status",
        "Paired Tx Time", "Paired Tx Type", "Paired Tx Amount", "Pairing Gap (Hours)", "Audit Week"
    ]
    
    # Combine Pyxis current week and previous week
    df_pyxis_filt['Audit Week'] = 'Current Week'
    if prev_pyxis_rows:
        df_prev_pyxis_all = pd.concat(prev_pyxis_rows, ignore_index=True)
        # Align columns
        for col in df_pyxis_filt.columns:
            if col not in df_prev_pyxis_all.columns:
                df_prev_pyxis_all[col] = None
        df_prev_pyxis_all = df_prev_pyxis_all[df_pyxis_filt.columns]
        df_pyxis_filt_combined = pd.concat([df_pyxis_filt, df_prev_pyxis_all], ignore_index=True)
    else:
        df_pyxis_filt_combined = df_pyxis_filt.copy()
        
    df_raw_out = df_pyxis_filt_combined[[
        'OrderNumber_Clean', 'PatientID_Clean', 'PatientName_Clean', 'TransactionDateTime_Clean',
        'TransactionType_Clean', 'TransactionStatus', 'MedDescription', 'DispenseAmount', 'WasteAmount', 'ParsedConc',
        'ParsedConcUnit', 'ParsedTotalVol', 'CalculatedDose', 'CalculatedVol_ml', 'PairingStatus',
        'PairedTxTime', 'PairedTxType', 'PairedTxAmount', 'PairingGapHours', 'Audit Week'
    ]].rename(columns={
        'OrderNumber_Clean': 'OrderNumber',
        'PatientID_Clean': 'PatientID',
        'PatientName_Clean': 'PatientName',
        'TransactionDateTime_Clean': 'TransactionDateTime',
        'TransactionType_Clean': 'TransactionType',
        'TransactionStatus': 'Transaction Status'
    })
    
    def rule_pyxis(ws, r_idx, row_vals):
        p_status = str(row_vals[14]).strip()
        if p_status == "Matched":
            ws.cell(row=r_idx, column=15).fill = GREEN_FILL
            ws.cell(row=r_idx, column=15).font = GREEN_FONT
        elif p_status == "Unmatched - Missing Waste":
            ws.cell(row=r_idx, column=15).fill = RED_FILL
            ws.cell(row=r_idx, column=15).font = RED_FONT
        elif p_status == "Unmatched - Excess Waste":
            ws.cell(row=r_idx, column=15).fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
            ws.cell(row=r_idx, column=15).font = Font(name="Segoe UI", size=10, color="856404", bold=True)
            
        # Highlight Previous Week rows
        week_val = str(row_vals[19]).strip()
        if week_val == 'Previous Week':
            ws.cell(row=r_idx, column=20).fill = PatternFill(start_color="E6F4EA", end_color="E6F4EA", fill_type="solid")
            ws.cell(row=r_idx, column=20).font = Font(name="Segoe UI", size=10, color="137333", bold=True)
            
    format_sheet(ws_pyxis_sheet, raw_headers, df_raw_out, rule_pyxis)
    
    # Reorder sheets to put worksheets in correct layout order
    sheet_order = [
        "Order Reconciliation Summary",
        "Epic Administrations Log",
        "Epic Flowsheet Log",
        "Paired Bags",
        "CS continuous infusions"
    ]
    wb._sheets = [wb[title] for title in sheet_order if title in wb.sheetnames]

    # Save file
    print(f"Saving workbook to {OUTPUT_PATH}...")
    try:
        wb.save(OUTPUT_PATH)
        print("Execution complete! Consolidated Step 3 report created successfully.")
    except PermissionError:
        locked_path = OUTPUT_PATH.replace(".xlsx", "_LOCKED.xlsx")
        print(f"\n[WARNING] Permission denied: {OUTPUT_PATH} is locked (likely open in Excel).")
        print(f"Saving a copy instead to: {locked_path}")
        wb.save(locked_path)
        print("Execution complete! Consolidated Step 3 report created successfully (saved as copy).")
        
    # Also save a copy to the default output path (overwriting the previous default file)
    default_out_path = os.path.join(BASE_DIR, "Output", "Step1_Pyxis_Analysis.xlsx")
    print(f"Saving copy to default path: {default_out_path}...")
    try:
        wb.save(default_out_path)
    except PermissionError:
        locked_default = default_out_path.replace(".xlsx", "_LOCKED.xlsx")
        print(f"[WARNING] Default path is locked. Saving a copy to: {locked_default}")
        wb.save(locked_default)
        
    # Also save a copy to the final report path with current date timestamp (YYYYMMDD)
    current_date_str = datetime.now().strftime("%Y%m%d")
    dated_report_path = os.path.join(BASE_DIR, "Output", f"Infusion_Reconciliation_Report_{current_date_str}.xlsx")
    print(f"Saving copy to dated final report path: {dated_report_path}...")
    try:
        wb.save(dated_report_path)
    except PermissionError:
        locked_dated = dated_report_path.replace(".xlsx", "_LOCKED.xlsx")
        print(f"[WARNING] Dated final report path is locked. Saving a copy to: {locked_dated}")
        wb.save(locked_dated)
        
    final_report_path = os.path.join(BASE_DIR, "Output", "Infusion_Reconciliation_Report.xlsx")
    print(f"Saving copy to final report path: {final_report_path}...")
    try:
        wb.save(final_report_path)
    except PermissionError:
        locked_final = final_report_path.replace(".xlsx", "_LOCKED.xlsx")
        print(f"[WARNING] Final report path is locked. Saving a copy to: {locked_final}")
        wb.save(locked_final)
    print("=" * 60)

def main():
    import glob
    print("=" * 60)
    print("STEP 3: EPIC FLOWSHEET INTEGRATION & RECONCILIATION")
    print("=" * 60)
    
    # Locate all CS continuous infusions weekly reports
    input_dir = os.path.join(BASE_DIR, "Input")
    csv_files = [os.path.join(input_dir, f) for f in os.listdir(input_dir)
                 if f.lower().startswith("cs continuous infusions - weekly report") and f.lower().endswith(".csv")]
    if not csv_files:
        csv_files = [os.path.join(input_dir, f) for f in os.listdir(input_dir)
                     if f.lower().endswith(".csv") and 'audittransaction' not in f.lower() and 'flowsheet' not in f.lower() and 'volume' not in f.lower()]
    if not csv_files:
        csv_files = [PYXIS_CSV_PATH]
        
    # Sort files by date ascending so that the latest week is processed last (and thus overwrites the default report)
    def get_pyxis_file_date(filepath):
        filename = os.path.basename(filepath)
        match = re.search(r"weekly report\s+([\d\-]+)", filename)
        if match:
            try:
                return datetime.strptime(match.group(1), "%m-%d-%Y")
            except:
                pass
        return datetime.min
    csv_files.sort(key=get_pyxis_file_date)
        
    # Select the latest weekly report as the current week Pyxis file to process
    if csv_files:
        latest_pyxis = csv_files[-1]
        print(f"Selected Latest Weekly Pyxis File (Current Week): {os.path.basename(latest_pyxis)}")
        prev_pyxis = None
        if len(csv_files) >= 2:
            prev_pyxis = csv_files[-2]
            print(f"Selected Previous Weekly Pyxis File (Previous Week): {os.path.basename(prev_pyxis)}")
        try:
            process_pyxis_file(latest_pyxis, prev_pyxis_path=prev_pyxis)
        except Exception as e:
            print(f"\n[ERROR] Failed processing Pyxis file: {os.path.basename(latest_pyxis)}")
            print(f"Reason: {e}")
            import traceback
            traceback.print_exc()
    else:
        print("[ERROR] No Pyxis CSV report files found to process.")
            
    print("\n" + "=" * 60)
    print("WEEKLY PYXIS REPORT RUN COMPLETE!")
    print("=" * 60)

if __name__ == "__main__":
    main()
