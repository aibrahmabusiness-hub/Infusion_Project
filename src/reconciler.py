import os
import json
import math
import pandas as pd
from datetime import datetime
from src.config import (
    PYXIS_PATH, EPIC_ADMIN_PATH, EPIC_FLOWSHEET_PATH, HISTORICAL_STATE_PATH,
    VEND_WASTE_MATCH_WINDOW_HOURS, MAX_HANG_TIME_HOURS, AI_MATCH_CONFIDENCE_THRESHOLD
)
from src.parser import parse_med_description, parse_amount_string, convert_mass_to_volume, convert_volume_to_mass
from src.matcher import match_medications

def load_historical_state():
    """
    Loads active orders from previous weeks.
    """
    if os.path.exists(HISTORICAL_STATE_PATH):
        try:
            with open(HISTORICAL_STATE_PATH, 'r') as f:
                return json.load(f)
        except Exception as e:
            print(f"Warning: Failed to load historical state: {e}. Starting fresh.")
    return {}

def save_historical_state(state):
    """
    Saves active orders to historical state file.
    """
    try:
        with open(HISTORICAL_STATE_PATH, 'w') as f:
            json.dump(state, f, indent=4)
    except Exception as e:
        print(f"Error saving historical state: {e}")

def safe_to_datetime(val):
    if pd.isna(val):
        return None
    if isinstance(val, datetime):
        return val
    try:
        return pd.to_datetime(str(val))
    except Exception:
        return None

def clean_id(val):
    if pd.isna(val):
        return ""
    # Remove .0 if it's float represented as string
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s

def clean_patient_name(name):
    if pd.isna(name) or not isinstance(name, str):
        return ""
    # Convert "Last, First" or "First Last" to normalized lowercase alphanumeric string
    name_clean = "".join(c for c in name.lower() if c.isalnum())
    return name_clean

def filter_by_vend_windows(df_to_filter, p_sub, time_column, time_window_hours_before=1.0, max_hang_time_hours=96.0):
    if p_sub.empty or df_to_filter.empty:
        return df_to_filter
        
    p_times = p_sub['TransactionDateTime_Clean'].dropna().tolist()
    if not p_times:
        return df_to_filter
        
    p_times = sorted(p_times)
    start_w = p_times[0] - pd.Timedelta(hours=time_window_hours_before)
    end_w = p_times[-1] + pd.Timedelta(hours=max_hang_time_hours)
    
    times = pd.to_datetime(df_to_filter[time_column], errors='coerce')
    mask = (times >= start_w) & (times <= end_w)
    return df_to_filter[mask]

CONTROLLED_DRUGS = ['fentanyl', 'midazolam', 'versed', 'ketamine', 'morphine', 'hydromorphone', 'dilaudid']

def is_controlled(text):
    if not isinstance(text, str):
        return False
    text_lower = text.lower()
    return any(drug in text_lower for drug in CONTROLLED_DRUGS)

def run_reconciliation():
    """
    Processes Pyxis, Epic Administrations, and Epic Flowsheet records.
    Performs matching, pairing, calculations, historical accumulation, and exception generation.
    Returns a dictionary of dataframes/lists representing the results.
    """
    # Load Data
    print("Loading data files...")
    df_pyxis_raw = pd.read_csv(PYXIS_PATH)
    if os.path.exists(EPIC_ADMIN_PATH):
        df_admin_raw = pd.read_excel(EPIC_ADMIN_PATH)
    else:
        print(f"[WARNING] Epic Administrations report not found at: {EPIC_ADMIN_PATH}. Running in flowsheet-only mode.")
        df_admin_raw = pd.DataFrame(columns=['Order ID', 'MRN', 'Patient', 'Medication', 'AdministrationAction', 'AdministrationInstant', 'Dose', 'DoseUnit', 'WeightAtRelease_X'])
    df_flow_raw = pd.read_excel(EPIC_FLOWSHEET_PATH)
    
    # Load historical state
    historical_state = load_historical_state()
    
    # Preprocess Data
    # Pyxis
    df_pyxis = df_pyxis_raw.copy()
    df_pyxis['OrderNumber_Clean'] = df_pyxis['OrderNumber'].apply(clean_id)
    df_pyxis['PatientID_Clean'] = df_pyxis['PatientID'].apply(clean_id)
    df_pyxis['PatientName_Norm'] = df_pyxis['PatientName'].apply(clean_patient_name)
    df_pyxis['TransactionDateTime_Clean'] = df_pyxis['TransactionDateTime'].apply(safe_to_datetime)
    
    # Filter Pyxis to controlled substances and keep all relevant transaction types (Vend, Waste, Stock Return, Vend (Cancel))
    df_pyxis = df_pyxis[df_pyxis['MedDescription'].apply(is_controlled)].copy()
    df_pyxis['TransactionType_Clean'] = df_pyxis['TransactionType'].astype(str).str.strip()
    df_pyxis = df_pyxis[df_pyxis['TransactionType_Clean'].isin(['Vend', 'Waste', 'Stock Return', 'Vend (Cancel)'])].copy()
    
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
        
    df_pyxis['TransactionStatus'] = df_pyxis['TransactionType_Clean'].apply(determine_status)
    
    df_admin = df_admin_raw.copy()
    df_admin['Order ID_Clean'] = df_admin['Order ID'].apply(clean_id)
    df_admin['MRN_Clean'] = df_admin['MRN'].apply(clean_id)
    df_admin['Patient_Norm'] = df_admin['Patient'].apply(clean_patient_name)
    df_admin['AdministrationInstant_Clean'] = df_admin['AdministrationInstant'].apply(safe_to_datetime)
    if not df_admin.empty:
        df_admin = df_admin[df_admin['Medication'].apply(is_controlled)].copy()
    
    # Epic Flowsheet - Filter to controlled substances
    df_flow = df_flow_raw.copy()
    df_flow['Order Med ID_Clean'] = df_flow['Order Med ID'].apply(clean_id)
    df_flow['MRN_Clean'] = df_flow['MRN'].apply(clean_id)
    df_flow['Patient Name_Norm'] = df_flow['Patient Name'].apply(clean_patient_name)
    df_flow['Recorded Time_Clean'] = df_flow['Recorded Time'].apply(safe_to_datetime)
    df_flow['Flowsheet Value'] = pd.to_numeric(df_flow['Flowsheet Value'], errors='coerce')
    if not df_flow.empty:
        df_flow = df_flow[df_flow['Flowsheet Row Name'].apply(is_controlled)].copy()
    
    print(f"Loaded and filtered: {len(df_pyxis)} Pyxis transactions, {len(df_admin)} Epic Admin records, and {len(df_flow)} Flowsheet entries.")
    
    # Identify unique Order IDs across files for auditing infusions
    pyxis_orders = set(df_pyxis['OrderNumber_Clean'].dropna().unique()) - {''}
    flow_orders = set(df_flow['Order Med ID_Clean'].dropna().unique()) - {''}
    
    # Audit scope: Union of Pyxis dispensing and Flowsheet infusion volumes
    all_orders = pyxis_orders.union(flow_orders)
    
    exceptions = []
    reconciliation_rows = []
    administered_calc_rows = []
    pyxis_details = []
    epic_admin_details = []
    epic_flowsheet_details = []
    
    # Group Pyxis Vends and Wastes for pairing
    # We will pair them BEFORE performing order-level reconciliation
    print("Performing Vend/Waste pairing in Pyxis...")
    df_pyxis['PairingStatus'] = 'Unmatched'
    # For Stock Return and Cancelled rows, set PairingStatus to Excluded directly
    df_pyxis.loc[df_pyxis['TransactionStatus'].isin(['Stock Return', 'Cancelled']), 'PairingStatus'] = 'Excluded'
    df_pyxis['PairedTxTime'] = None
    df_pyxis['PairedTxType'] = None
    df_pyxis['PairedTxAmount'] = None
    df_pyxis['PairingGapHours'] = None
    df_pyxis['ParsedConc_per_ml'] = None
    df_pyxis['ParsedConc_unit'] = None
    df_pyxis['ParsedTotalVol_ml'] = None
    df_pyxis['CalculatedDose'] = None
    df_pyxis['CalculatedVol_ml'] = None
    
    # For tracking unique row identifier in Pyxis
    df_pyxis['TxIndex'] = range(len(df_pyxis))
    
    # Sort Pyxis by time
    df_pyxis = df_pyxis.sort_values('TransactionDateTime_Clean')
    
    for (order_num, med_desc), group in df_pyxis.groupby(['OrderNumber_Clean', 'MedDescription']):
        if not order_num or pd.isna(med_desc):
            continue
            
        # Parse concentration details for this group
        parsed_med = parse_med_description(med_desc)
        conc_per_ml = parsed_med['concentration_per_ml']
        conc_unit = parsed_med['concentration_unit']
        tot_vol = parsed_med['total_volume']
        
        vends = group[(group['TransactionType'] == 'Vend') & (group['TransactionStatus'] == 'Active')].copy()
        wastes = group[(group['TransactionType'] == 'Waste') & (group['TransactionStatus'] == 'Active')].copy()
        
        # Populate parsed columns in group
        for idx in group.index:
            tx_type = group.loc[idx, 'TransactionType']
            raw_disp = group.loc[idx, 'DispenseAmount']
            raw_waste = group.loc[idx, 'WasteAmount']
            
            df_pyxis.loc[idx, 'ParsedConc_per_ml'] = conc_per_ml
            df_pyxis.loc[idx, 'ParsedConc_unit'] = conc_unit
            df_pyxis.loc[idx, 'ParsedTotalVol_ml'] = tot_vol
            
            if tx_type == 'Vend':
                val, unit = parse_amount_string(raw_disp)
                if val is not None:
                    df_pyxis.loc[idx, 'CalculatedDose'] = val
                    df_pyxis.loc[idx, 'CalculatedVol_ml'] = convert_mass_to_volume(val, unit, conc_per_ml, conc_unit)
            elif tx_type == 'Waste':
                val, unit = parse_amount_string(raw_waste)
                if val is not None:
                    df_pyxis.loc[idx, 'CalculatedDose'] = val
                    df_pyxis.loc[idx, 'CalculatedVol_ml'] = convert_mass_to_volume(val, unit, conc_per_ml, conc_unit)
                    
        # Match each waste with its closest unmatched Vend
        paired_vends = set() # store TxIndices of paired vends
        
        for w_idx, waste_row in wastes.iterrows():
            w_time = waste_row['TransactionDateTime_Clean']
            w_tx_idx = waste_row['TxIndex']
            raw_waste = waste_row['WasteAmount']
            
            # Find closest vend
            best_vend_idx = None
            min_diff_hours = float('inf')
            
            for v_idx, vend_row in vends.iterrows():
                v_tx_idx = vend_row['TxIndex']
                if v_tx_idx in paired_vends:
                    continue
                v_time = vend_row['TransactionDateTime_Clean']
                
                if w_time and v_time:
                    diff_hours = abs((w_time - v_time).total_seconds()) / 3600.0
                    if diff_hours < min_diff_hours:
                        min_diff_hours = diff_hours
                        best_vend_idx = v_idx
            
            # Pair by closest transaction datetime
            if best_vend_idx is not None:
                vend_row = vends.loc[best_vend_idx]
                v_tx_idx = vend_row['TxIndex']
                paired_vends.add(v_tx_idx)
                
                # Update Waste Row
                df_pyxis.loc[w_idx, 'PairingStatus'] = 'Matched'
                df_pyxis.loc[w_idx, 'PairedTxTime'] = vend_row['TransactionDateTime_Clean']
                df_pyxis.loc[w_idx, 'PairedTxType'] = 'Vend'
                df_pyxis.loc[w_idx, 'PairedTxAmount'] = vend_row['DispenseAmount']
                df_pyxis.loc[w_idx, 'PairingGapHours'] = min_diff_hours
                
                # Update Vend Row
                df_pyxis.loc[best_vend_idx, 'PairingStatus'] = 'Matched'
                df_pyxis.loc[best_vend_idx, 'PairedTxTime'] = w_time
                df_pyxis.loc[best_vend_idx, 'PairedTxType'] = 'Waste'
                df_pyxis.loc[best_vend_idx, 'PairedTxAmount'] = raw_waste
                df_pyxis.loc[best_vend_idx, 'PairingGapHours'] = min_diff_hours
            else:
                # Waste remains unmatched (Excess Waste)
                df_pyxis.loc[w_idx, 'PairingStatus'] = 'Unmatched - Excess Waste'
                if best_vend_idx is not None:
                    # Let's see if it's ambiguous
                    df_pyxis.loc[w_idx, 'PairingGapHours'] = min_diff_hours
                    
        # Update remaining unmatched Vends
        for v_idx, vend_row in vends.iterrows():
            v_tx_idx = vend_row['TxIndex']
            if v_tx_idx not in paired_vends:
                df_pyxis.loc[v_idx, 'PairingStatus'] = 'Unmatched - Missing Waste'
                
    # 3. Order-Level Reconciliation
    print("Reconciling orders...")
    unpaired_order_ids = set(df_pyxis[df_pyxis['PairingStatus'].str.startswith('Unmatched', na=False)]['OrderNumber_Clean'].dropna().unique())
    for order_id in all_orders:
        if not order_id:
            continue
            
        # Get matching subsets
        p_sub = df_pyxis[df_pyxis['OrderNumber_Clean'] == order_id]
        a_sub = df_admin[df_admin['Order ID_Clean'] == order_id]
        f_sub = df_flow[df_flow['Order Med ID_Clean'] == order_id]
        
        # Apply window filtering to prevent including administrations/flowsheets from other weeks/bags
        a_sub = filter_by_vend_windows(a_sub, p_sub, 'AdministrationInstant_Clean')
        f_sub = filter_by_vend_windows(f_sub, p_sub, 'Recorded Time_Clean')
        
        patient_name = "Unknown"
        mrn = "Unknown"
        med_pyxis = "Unknown"
        med_epic = "Unknown"
        flow_row_name = "Unknown"
        
        # Determine Patient Name and MRN
        if not p_sub.empty:
            patient_name = p_sub.iloc[0]['PatientName']
            mrn = clean_id(p_sub.iloc[0]['PatientID'])
            med_pyxis = p_sub.iloc[0]['MedDescription']
        elif not a_sub.empty:
            patient_name = a_sub.iloc[0]['Patient']
            mrn = a_sub.iloc[0]['MRN_Clean']
            med_epic = a_sub.iloc[0]['Medication']
        elif not f_sub.empty:
            patient_name = f_sub.iloc[0]['Patient Name']
            mrn = f_sub.iloc[0]['MRN_Clean']
            flow_row_name = f_sub.iloc[0]['Flowsheet Row Name']
            
        # Standardize MRN/Patient across fields
        if not a_sub.empty and mrn == "Unknown":
            mrn = a_sub.iloc[0]['MRN_Clean']
        if not f_sub.empty and mrn == "Unknown":
            mrn = f_sub.iloc[0]['MRN_Clean']
            
        # Extract med names if not filled
        if not a_sub.empty and med_epic == "Unknown":
            med_epic = a_sub.iloc[0]['Medication']
        if not f_sub.empty and flow_row_name == "Unknown":
            flow_row_name = f_sub.iloc[0]['Flowsheet Row Name']
            
        # Check matching across files
        medication_matched = True
        match_confidence = 1.0
        match_method = "Deterministic Match"
        match_explanation = "Same order association"
        
        # If we have Pyxis and Epic Admin
        if not p_sub.empty and not a_sub.empty:
            medication_matched, match_confidence, match_method, match_explanation = match_medications(
                med_pyxis, med_epic, threshold=AI_MATCH_CONFIDENCE_THRESHOLD
            )
        # If we have Pyxis and Flowsheet
        elif not p_sub.empty and not f_sub.empty:
            medication_matched, match_confidence, match_method, match_explanation = match_medications(
                med_pyxis, flow_row_name, threshold=AI_MATCH_CONFIDENCE_THRESHOLD
            )
        # If we have Epic Admin and Flowsheet
        elif not a_sub.empty and not f_sub.empty:
            medication_matched, match_confidence, match_method, match_explanation = match_medications(
                med_epic, flow_row_name, threshold=AI_MATCH_CONFIDENCE_THRESHOLD
            )
            
        # Parse Pyxis details if available
        conc_per_ml = None
        conc_unit = None
        
        if not p_sub.empty:
            p_desc = p_sub.iloc[0]['MedDescription']
            parsed_med = parse_med_description(p_desc)
            conc_per_ml = parsed_med['concentration_per_ml']
            conc_unit = parsed_med['concentration_unit']
            
        # Check Administration Actions to determine reconciliation unit and look for Stopped/Restarted actions
        reconciliation_unit = 'mL'  # Default
        has_bolus = False
        has_stopped = False
        has_restarted = False
        admin_actions = []
        
        if not a_sub.empty:
            actions = a_sub['AdministrationAction'].dropna().unique()
            admin_actions = list(actions)
            if any('bolus' in act.lower() for act in actions):
                has_bolus = True
                reconciliation_unit = conc_unit if conc_unit else 'mg'
            if any('stop' in act.lower() for act in actions):
                has_stopped = True
            if any('restart' in act.lower() for act in actions):
                has_restarted = True
                
        # Calculate Dispensed, Administered, and Waste amounts in both units (mg/mcg and mL)
        # 1. Pyxis Dispensed (Active Vends only)
        disp_vends = p_sub[(p_sub['TransactionType'] == 'Vend') & (p_sub['TransactionStatus'] == 'Active')]
        total_disp_dose = 0.0
        total_disp_vol = 0.0
        
        for _, row in disp_vends.iterrows():
            dose = row['CalculatedDose']
            vol = row['CalculatedVol_ml']
            if dose is not None:
                total_disp_dose += dose
            if vol is not None:
                total_disp_vol += vol
                
        # 2. Pyxis Documented Waste (Active Wastes only)
        disp_wastes = p_sub[(p_sub['TransactionType'] == 'Waste') & (p_sub['TransactionStatus'] == 'Active')]
        total_waste_dose = 0.0
        total_waste_vol = 0.0
        
        for _, row in disp_wastes.iterrows():
            dose = row['CalculatedDose']
            vol = row['CalculatedVol_ml']
            if dose is not None:
                total_waste_dose += dose
            if vol is not None:
                total_waste_vol += vol
                
        # 3. Epic Flowsheet Administered (documented in mL)
        total_admin_vol = f_sub['Flowsheet Value'].sum()
        total_admin_dose = 0.0
        if conc_per_ml is not None:
            total_admin_dose = total_admin_vol * conc_per_ml
            
        # Cross-Week Accumulation Logic
        # Check if we have historical data for this order
        first_seen_str = None
        last_seen_str = None
        
        # Get timestamps from the current files
        all_timestamps = []
        if not p_sub.empty:
            all_timestamps.extend(p_sub['TransactionDateTime_Clean'].dropna().tolist())
        if not a_sub.empty:
            all_timestamps.extend(a_sub['AdministrationInstant_Clean'].dropna().tolist())
        if not f_sub.empty:
            all_timestamps.extend(f_sub['Recorded Time_Clean'].dropna().tolist())
            
        if all_timestamps:
            min_ts = min(all_timestamps)
            max_ts = max(all_timestamps)
            first_seen_str = min_ts.isoformat()
            last_seen_str = max_ts.isoformat()
            
        acc_dispensed_dose = total_disp_dose
        acc_dispensed_vol = total_disp_vol
        acc_administered_vol = total_admin_vol
        acc_administered_dose = total_admin_dose
        acc_waste_dose = total_waste_dose
        acc_waste_vol = total_waste_vol
        
        is_new_order = True
        
        if order_id in historical_state:
            is_new_order = False
            hist = historical_state[order_id]
            # Accumulate
            acc_dispensed_dose += hist.get('accumulated_dispensed_dose', 0.0)
            acc_dispensed_vol += hist.get('accumulated_dispensed_vol', 0.0)
            acc_administered_vol += hist.get('accumulated_administered_vol', 0.0)
            acc_administered_dose += hist.get('accumulated_administered_dose', 0.0)
            acc_waste_dose += hist.get('accumulated_waste_dose', 0.0)
            acc_waste_vol += hist.get('accumulated_waste_vol', 0.0)
            
            # Retain original first seen
            first_seen_str = hist.get('first_seen_time', first_seen_str)
            
        # Calculate expected waste and variance on cumulative totals
        if reconciliation_unit == 'mL':
            dispensed_rec = acc_dispensed_vol
            administered_rec = acc_administered_vol
            waste_rec = acc_waste_vol
        else:
            dispensed_rec = acc_dispensed_dose
            administered_rec = acc_administered_dose
            waste_rec = acc_waste_dose
            
        expected_waste_rec = dispensed_rec - administered_rec
        variance_rec = expected_waste_rec - waste_rec
        
        # Calculate new fields
        total_bags = len(p_sub[(p_sub['TransactionType'] == 'Vend') & (p_sub['TransactionStatus'] == 'Active')])
        infusion_type = "Bolus" if has_bolus else "Continuous Infusion"
        variance_pct = abs(variance_rec) / dispensed_rec if dispensed_rec > 0.0 else 0.0
        
        # Check hang time exception
        hang_time_exceeded = False
        duration_hours = 0.0
        if first_seen_str and last_seen_str:
            try:
                first_dt = datetime.fromisoformat(first_seen_str)
                last_dt = datetime.fromisoformat(last_seen_str)
                duration_hours = (last_dt - first_dt).total_seconds() / 3600.0
                if duration_hours > MAX_HANG_TIME_HOURS:
                    hang_time_exceeded = True
            except Exception:
                pass
                
        # Determine Status and Exception Flags
        flags = []
        status = "Reconciled"
        
        # 1. Stopped / Restarted Action Exception
        if has_stopped:
            flags.append("Stopped Action")
            status = "Exception"
            exceptions.append({
                'OrderID': order_id, 'MRN': mrn, 'PatientName': patient_name,
                'SourceFile': 'Epic Medication Administrations', 'ExceptionType': 'Stopped actions',
                'RecordIdentifier': 'AdministrationAction', 'Description': f"Order contains a 'Stopped' action in Epic: {admin_actions}",
                'ActionNeeded': 'Verify the stopped status and manually reconcile remaining medication.'
            })
        if has_restarted:
            flags.append("Restarted Action")
            status = "Exception"
            exceptions.append({
                'OrderID': order_id, 'MRN': mrn, 'PatientName': patient_name,
                'SourceFile': 'Epic Medication Administrations', 'ExceptionType': 'Restarted actions',
                'RecordIdentifier': 'AdministrationAction', 'Description': f"Order contains a 'Restarted' action in Epic: {admin_actions}",
                'ActionNeeded': 'Verify if restarted action caused duplicate flowsheet lines or interrupted workflows.'
            })
            
        # 2. Naming / Medication Match Exception
        if not medication_matched:
            flags.append("Medication Match Failure")
            status = "Exception"
            exceptions.append({
                'OrderID': order_id, 'MRN': mrn, 'PatientName': patient_name,
                'SourceFile': 'Pyxis / Epic Naming Matcher', 'ExceptionType': 'Ambiguous medication matching',
                'RecordIdentifier': 'Medication Naming', 'Description': f"Could not confidently match Pyxis '{med_pyxis}', Epic '{med_epic}', and Flowsheet '{flow_row_name}'. {match_explanation}",
                'ActionNeeded': 'Manually verify if the flowsheet row name matches the dispensed medication.'
            })
            
        # 3. Missing Pyxis Activity Exception
        if p_sub.empty and (not a_sub.empty or not f_sub.empty):
            flags.append("Missing Pyxis Activity")
            status = "Exception"
            exceptions.append({
                'OrderID': order_id, 'MRN': mrn, 'PatientName': patient_name,
                'SourceFile': 'Pyxis Audit Transaction', 'ExceptionType': 'Missing Pyxis activity',
                'RecordIdentifier': 'Pyxis CSV', 'Description': "Epic administration/flowsheet records exist but no Pyxis dispensed or waste transactions found.",
                'ActionNeeded': 'Confirm if medication was dispensed under a different order number or override.'
            })
            
        # 4. Missing Administration Activity Exception
        if not p_sub.empty and f_sub.empty:
            flags.append("Missing Administration Activity")
            status = "Exception"
            exceptions.append({
                'OrderID': order_id, 'MRN': mrn, 'PatientName': patient_name,
                'SourceFile': 'Epic Flowsheet', 'ExceptionType': 'Missing administration activity',
                'RecordIdentifier': 'Epic Flowsheet', 'Description': "Dispensed from Pyxis but no volumes found in Epic Flowsheet.",
                'ActionNeeded': 'Check if clinician failed to document infusion volume or if patient did not receive infusion.'
            })
            
        # 5. Missing Waste Documentation Exception
        # If there are unmatched vends in this order
        unmatched_vends = p_sub[p_sub['PairingStatus'] == 'Unmatched - Missing Waste']
        if not unmatched_vends.empty:
            flags.append("Missing Waste Documentation")
            status = "Exception"
            for _, row in unmatched_vends.iterrows():
                exceptions.append({
                    'OrderID': order_id, 'MRN': mrn, 'PatientName': patient_name,
                    'SourceFile': 'Pyxis Audit Transaction', 'ExceptionType': 'Missing waste documentation',
                    'RecordIdentifier': f"Vend at {row['TransactionDateTime']}", 'Description': f"Vend transaction ({row['DispenseAmount']}) has no paired Waste transaction within 8 hours.",
                    'ActionNeeded': 'Verify if clinician wasted remaining medication or forgot to document waste in Pyxis.'
                })
                
        # 6. Excess Waste Documentation Exception
        # If there are unmatched wastes in this order
        unmatched_wastes = p_sub[p_sub['PairingStatus'] == 'Unmatched - Excess Waste']
        if not unmatched_wastes.empty:
            flags.append("Excess Waste Documentation")
            status = "Exception"
            for _, row in unmatched_wastes.iterrows():
                exceptions.append({
                    'OrderID': order_id, 'MRN': mrn, 'PatientName': patient_name,
                    'SourceFile': 'Pyxis Audit Transaction', 'ExceptionType': 'Excess waste documentation',
                    'RecordIdentifier': f"Waste at {row['TransactionDateTime']}", 'Description': f"Waste transaction ({row['WasteAmount']}) has no paired Vend transaction within 8 hours.",
                    'ActionNeeded': 'Check if waste represents a bag dispensed in previous week, or if vend transaction was cancel/omitted.'
                })
                
        # 7. Bolus Concentration Conversion Exception
        if has_bolus and conc_per_ml is None:
            flags.append("Missing Conc for Bolus Conversion")
            status = "Exception"
            exceptions.append({
                'OrderID': order_id, 'MRN': mrn, 'PatientName': patient_name,
                'SourceFile': 'Pyxis MedDescription Parser', 'ExceptionType': 'Missing concentration data for Bolus conversions',
                'RecordIdentifier': 'Pyxis MedDescription', 'Description': f"Bolus conversion required but concentration could not be extracted from Pyxis medication description: '{med_pyxis}'",
                'ActionNeeded': 'Manually parse medication concentration and convert flowsheet mL to mg/mcg.'
            })
            
        # 8. Maximum Hang Time Exception
        if hang_time_exceeded:
            flags.append("Exceeded Max Hang Time (96h)")
            status = "Exception"
            exceptions.append({
                'OrderID': order_id, 'MRN': mrn, 'PatientName': patient_name,
                'SourceFile': 'Cross-Week Order Tracker', 'ExceptionType': 'Orders that remain unresolved after the maximum hang time',
                'RecordIdentifier': 'First seen to Last seen duration', 'Description': f"Order activity spans {duration_hours:.1f} hours, exceeding the 96-hour maximum hang time.",
                'ActionNeeded': 'Review clinical necessity of long-running infusion and check for delayed documentation.'
            })
            
        # 9. Discrepancy / Variance checking
        # If not already an exception, check if variance is non-zero (greater than a small epsilon to handle floats)
        is_discrepancy = False
        if status != "Exception":
            if abs(variance_rec) > 1e-5:
                status = "Discrepancy"
                is_discrepancy = True
                flags.append("Material Variance")
                exceptions.append({
                    'OrderID': order_id, 'MRN': mrn, 'PatientName': patient_name,
                    'SourceFile': 'Reconciliation Logic', 'ExceptionType': 'Discrepancy',
                    'RecordIdentifier': f"Variance: {variance_rec:.2f} {reconciliation_unit}", 'Description': f"Reconciliation variance is {variance_rec:.2f} {reconciliation_unit} (Expected Waste: {expected_waste_rec:.2f}, Documented Waste: {waste_rec:.2f})",
                    'ActionNeeded': 'Investigate discrepancy between expected waste (Dispensed - Administered) and actual documented waste.'
                })
                
        has_unmatched = "Yes" if order_id in unpaired_order_ids else "No"
        admin_status = "Missing Admin Charting" if administered_rec <= 0.0 else "Charted"
        
        # Get unique list of charted actions
        admin_actions_str = "N/A"
        if not a_sub.empty:
            seen_act = set()
            unique_actions = [a for a in a_sub['AdministrationAction'].astype(str).str.strip().tolist() if not (a in seen_act or seen_act.add(a))]
            admin_actions_str = ", ".join(unique_actions)
            
        # Store detailed reconciliation row
        reconciliation_rows.append({
            'OrderID': order_id,
            'PatientName': patient_name,
            'Medication': med_pyxis if med_pyxis != "Unknown" else (med_epic if med_epic != "Unknown" else flow_row_name),
            'ReconciliationUnit': reconciliation_unit,
            'TotalDispensed': dispensed_rec,
            'TotalWaste': waste_rec,
            'AvailableToAdminister': dispensed_rec - waste_rec,
            'TotalAdministered': administered_rec,
            'AdditionalInWaste': variance_rec,
            'OrderStatus': status,
            'HasUnmatchedCabinetPairs': has_unmatched,
            'CumulativeAdministeredStatus': admin_status,
            'AdministrationAction': admin_actions_str
        })
        
        administered_calc_rows.append({
            'OrderID': order_id,
            'PatientName': patient_name,
            'Medication': med_pyxis if med_pyxis != "Unknown" else (med_epic if med_epic != "Unknown" else flow_row_name),
            'ReconciliationUnit': reconciliation_unit,
            'FlowsheetEntriesCount': len(f_sub),
            'TotalVolume_mL': total_admin_vol,
            'ConcentrationPerML': conc_per_ml if conc_per_ml is not None else 0.0,
            'ConvertedDose': total_admin_dose,
            'FinalCumulativeAdministered': administered_rec
        })
        
        # Save active/incomplete state
        # An order is completed if variance is 0 AND no exceptions AND the order isn't ongoing
        # Since we cannot check if the order is ongoing, we consider it completed if it reconciles to 0.
        # If it doesn't reconcile, we keep it in state so we accumulate more next week (up to 96h).
        is_completed = (status == "Reconciled")
        
        if not is_completed and duration_hours <= MAX_HANG_TIME_HOURS:
            historical_state[order_id] = {
                'accumulated_dispensed_dose': acc_dispensed_dose,
                'accumulated_dispensed_vol': acc_dispensed_vol,
                'accumulated_administered_vol': acc_administered_vol,
                'accumulated_administered_dose': acc_administered_dose,
                'accumulated_waste_dose': acc_waste_dose,
                'accumulated_waste_vol': acc_waste_vol,
                'first_seen_time': first_seen_str,
                'last_seen_time': last_seen_str,
                'status': 'active'
            }
        elif order_id in historical_state:
            # Remove completed or outdated order from state
            del historical_state[order_id]
            
        # Append sub-records for detailed sheets
        pyxis_details.append(p_sub)
        epic_admin_details.append(a_sub)
        epic_flowsheet_details.append(f_sub)
        
    # Save the updated historical state
    save_historical_state(historical_state)
    
    # Consolidate detail sheets
    df_reconciliation = pd.DataFrame(reconciliation_rows)
    df_administered_summary = pd.DataFrame(administered_calc_rows)
    df_pyxis_out = pd.concat(pyxis_details) if pyxis_details else pd.DataFrame(columns=df_pyxis.columns)
    df_admin_out = pd.concat(epic_admin_details) if epic_admin_details else pd.DataFrame(columns=df_admin.columns)
    df_flow_out = pd.concat(epic_flowsheet_details) if epic_flowsheet_details else pd.DataFrame(columns=df_flow.columns)
    df_exceptions = pd.DataFrame(exceptions) if exceptions else pd.DataFrame(columns=['OrderID', 'MRN', 'PatientName', 'SourceFile', 'ExceptionType', 'RecordIdentifier', 'Description', 'ActionNeeded'])
    
    # Remove cleanup columns from details to keep output clean
    pyxis_cols_to_keep = [
        'OrderNumber', 'PatientID', 'PatientName', 'TransactionDateTime', 'TransactionType',
        'TransactionStatus', 'MedDescription', 'DispenseAmount', 'WasteAmount',
        'ParsedConc_per_ml', 'ParsedConc_unit', 'ParsedTotalVol_ml', 'CalculatedDose', 'CalculatedVol_ml',
        'PairedTxTime', 'PairedTxType', 'PairedTxAmount', 'PairingStatus', 'PairingGapHours'
    ]
    df_pyxis_out = df_pyxis_out[[c for c in pyxis_cols_to_keep if c in df_pyxis_out.columns]]
    
    admin_cols_to_keep = [
        'Order ID', 'MRN', 'Patient', 'Medication', 'AdministrationAction', 'AdministrationInstant',
        'Dose', 'DoseUnit', 'WeightAtRelease_X'
    ]
    df_admin_out = df_admin_out[[c for c in admin_cols_to_keep if c in df_admin_out.columns]]
    
    flow_cols_to_keep = [
        'CSN', 'MRN', 'Patient Name', 'Order Med ID', 'Flowsheet Row Name', 'Flowsheet Display',
        'Recorded Time', 'Flowsheet Value'
    ]
    df_flow_out = df_flow_out[[c for c in flow_cols_to_keep if c in df_flow_out.columns]]
    
    return {
        'reconciliation': df_reconciliation,
        'administered_details': df_administered_summary,
        'pyxis_details': df_pyxis_out,
        'epic_admin_details': df_admin_out,
        'epic_flowsheet_details': df_flow_out,
        'exceptions': df_exceptions
    }
