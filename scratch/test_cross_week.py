import os
import glob
import re
import pandas as pd
from datetime import datetime

# Configurations
BASE_DIR = r"c:\Users\GaneshBhat\OneDrive - Novatio Solutions\Desktop\Infusion"
INPUT_DIR = os.path.join(BASE_DIR, "Input")

def parse_dates_from_filename(filename):
    match = re.search(r"(\d{8})_(\d{8})", filename)
    if match:
        try:
            start_dt = datetime.strptime(match.group(1), "%Y%m%d")
            end_dt = datetime.strptime(match.group(2), "%Y%m%d")
            return start_dt, end_dt
        except Exception:
            pass
    return None, None

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

def test_run():
    # 1. Find CSV
    csv_files = [f for f in glob.glob(os.path.join(INPUT_DIR, "*")) if f.lower().endswith('.csv') and 'audittransaction' not in os.path.basename(f).lower()]
    if not csv_files:
        csv_files = [f for f in glob.glob(os.path.join(INPUT_DIR, "*")) if f.lower().endswith('.csv')]
    if not csv_files:
        print("No CSV found!")
        return
        
    pyxis_path = csv_files[0]
    print(f"Loading Pyxis CSV from: {pyxis_path}")
    df_pyxis_raw = pd.read_csv(pyxis_path, low_memory=False)
    df_pyxis_raw['TransactionDateTime_Clean'] = pd.to_datetime(df_pyxis_raw['TransactionDateTime'], errors='coerce')
    
    pyxis_min_time = df_pyxis_raw['TransactionDateTime_Clean'].min()
    pyxis_max_time = df_pyxis_raw['TransactionDateTime_Clean'].max()
    print(f"Pyxis start date: {pyxis_min_time}, end date: {pyxis_max_time}")
    
    excel_files = glob.glob(os.path.join(INPUT_DIR, "*.xlsx")) + glob.glob(os.path.join(INPUT_DIR, "*.xls"))
    
    this_week_flow = None
    prev_week_flow = None
    this_week_admin = None
    prev_week_admin = None
    
    for f in excel_files:
        name = os.path.basename(f)
        start_dt, end_dt = parse_dates_from_filename(name)
        if start_dt and end_dt:
            if start_dt <= pyxis_min_time <= end_dt:
                if "flowsheet" in name.lower() or "volume" in name.lower():
                    this_week_flow = f
                elif "admin" in name.lower() or "action" in name.lower():
                    this_week_admin = f
            elif end_dt < pyxis_min_time and (pyxis_min_time - end_dt).days <= 3:
                if "flowsheet" in name.lower() or "volume" in name.lower():
                    prev_week_flow = f
                elif "admin" in name.lower() or "action" in name.lower():
                    prev_week_admin = f
                    
    print(f"This Week Flowsheet: {this_week_flow}")
    print(f"Previous Week Flowsheet: {prev_week_flow}")
    print(f"This Week Admin: {this_week_admin}")
    print(f"Previous Week Admin: {prev_week_admin}")

    # Load them
    df_flow_this = pd.read_excel(this_week_flow) if this_week_flow else pd.DataFrame()
    df_flow_prev = pd.read_excel(prev_week_flow) if prev_week_flow else pd.DataFrame()
    df_admin_this = pd.read_excel(this_week_admin) if this_week_admin else pd.DataFrame()
    df_admin_prev = pd.read_excel(prev_week_admin) if prev_week_admin else pd.DataFrame()
    
    # Process
    for df in [df_flow_this, df_flow_prev]:
        if not df.empty:
            df['OrderID_Clean'] = df['Order Med ID'].apply(clean_id)
            df['MRN_Clean'] = df['MRN'].apply(clean_id)
            df['PatientName_Norm'] = df['Patient Name'].apply(clean_patient_name)
            df['RecordedTime_Clean'] = pd.to_datetime(df['Recorded Time'], errors='coerce')
            df['FlowsheetValue_Num'] = pd.to_numeric(df['Flowsheet Value'], errors='coerce').fillna(0.0)
            
    for df in [df_admin_this, df_admin_prev]:
        if not df.empty:
            df['OrderID_Clean'] = df['Order ID'].apply(clean_id)
            df['MRN_Clean'] = df['MRN'].apply(clean_id)
            df['PatientName_Clean'] = df['Patient'].astype(str).str.strip()
            df['PatientName_Norm'] = df['Patient'].apply(clean_patient_name)
            df['AdministrationInstant_Clean'] = pd.to_datetime(df['AdministrationInstant'], errors='coerce')

    # Test for Order ID 6080119084
    order_id = '6080119084'
    df_pyxis_filt = df_pyxis_raw.copy()
    df_pyxis_filt['OrderNumber_Clean'] = df_pyxis_filt['OrderNumber'].apply(clean_id)
    df_pyxis_filt['TransactionDateTime_Clean'] = pd.to_datetime(df_pyxis_filt['TransactionDateTime'], errors='coerce')
    df_pyxis_filt['TransactionType_Clean'] = df_pyxis_filt['TransactionType'].astype(str).str.strip()
    df_pyxis_filt = df_pyxis_filt[df_pyxis_filt['TransactionType_Clean'].isin(['Vend', 'Waste', 'Stock Return', 'Vend (Cancel)'])].copy()
    
    p_sub = df_pyxis_filt[df_pyxis_filt['OrderNumber_Clean'] == order_id]
    
    f_this = df_flow_this[df_flow_this['OrderID_Clean'] == order_id] if not df_flow_this.empty else pd.DataFrame()
    f_prev = df_flow_prev[df_flow_prev['OrderID_Clean'] == order_id] if not df_flow_prev.empty else pd.DataFrame()
    
    e_this = df_admin_this[df_admin_this['OrderID_Clean'] == order_id] if not df_admin_this.empty else pd.DataFrame()
    e_prev = df_admin_prev[df_admin_prev['OrderID_Clean'] == order_id] if not df_admin_prev.empty else pd.DataFrame()
    
    if not p_sub.empty:
        p_times = sorted(p_sub['TransactionDateTime_Clean'].dropna().tolist())
        first_tx_time = p_times[0]
        last_tx_time = p_times[-1]
        
        first_tx_type = p_sub.sort_values('TransactionDateTime_Clean').iloc[0]['TransactionType_Clean']
        
        if first_tx_type == 'Waste':
            # Check e_prev or e_this for last dispense before first_tx_time
            e_all = pd.concat([e_prev, e_this]) if (not e_prev.empty or not e_this.empty) else pd.DataFrame()
            t_dispense = None
            if not e_all.empty:
                e_before = e_all[e_all['AdministrationInstant_Clean'] < first_tx_time].sort_values('AdministrationInstant_Clean', ascending=False)
                if not e_before.empty:
                    # Find first record with "new bag" or similar
                    for _, r in e_before.iterrows():
                        act_lower = str(r['AdministrationAction']).lower()
                        if 'bag' in act_lower or 'dose' in act_lower or 'start' in act_lower:
                            t_dispense = r['AdministrationInstant_Clean']
                            break
            if t_dispense:
                start_w = t_dispense - pd.Timedelta(hours=1)
                print(f"Found dispense in previous week's admin: {t_dispense}")
            else:
                start_w = first_tx_time - pd.Timedelta(hours=96)
        else:
            start_w = first_tx_time - pd.Timedelta(hours=1)
            
        end_w = last_tx_time + pd.Timedelta(hours=96)
        
        print(f"Calculation window: {start_w} to {end_w}")
        
        f_this_filt = f_this[(f_this['RecordedTime_Clean'] >= start_w) & (f_this['RecordedTime_Clean'] <= end_w)]
        f_prev_filt = f_prev[(f_prev['RecordedTime_Clean'] >= start_w) & (f_prev['RecordedTime_Clean'] <= end_w)]
        
        print(f"This week flowsheet records included: {len(f_this_filt)}")
        print(f"Previous week flowsheet records included: {len(f_prev_filt)}")

if __name__ == "__main__":
    test_run()
