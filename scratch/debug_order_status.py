import pandas as pd
import os
import shutil

BASE_DIR = r"c:\Users\GaneshBhat\OneDrive - Novatio Solutions\Desktop\Infusion"
PYXIS_CSV_PATH = os.path.join(BASE_DIR, "Input", "AuditTransactionDetail_RC.csv")
EPIC_ADMIN_PATH = os.path.join(BASE_DIR, "Input", "BCH_Administered_Medications_Report_20260617_1354.xlsx")
EPIC_FLOW_PATH = os.path.join(BASE_DIR, "Input", "BCH_Flowsheet_Med_Infusion_Volume_20260617_1342.xlsx")

def load_excel_safe(path):
    if not os.path.exists(path):
        return pd.DataFrame()
    temp_dir = os.path.dirname(path)
    temp_name = f"~temp_debug_{os.path.basename(path)}"
    temp_path = os.path.join(temp_dir, temp_name)
    try:
        shutil.copy(path, temp_path)
    except Exception:
        import subprocess
        subprocess.run(['powershell', '-Command', f'Copy-Item -Path "{path}" -Destination "{temp_path}" -Force'], shell=True)
    df = pd.read_excel(temp_path)
    try:
        os.remove(temp_path)
    except:
        pass
    return df

print("=== Debugging Order 6078156365 ===")
order_id = "6078156365"

# Load data
df_pyxis = pd.read_csv(PYXIS_CSV_PATH)
df_pyxis['OrderNumber_Clean'] = df_pyxis['OrderNumber'].astype(str).str.strip().str.replace(".0", "", regex=False)

df_epic = load_excel_safe(EPIC_ADMIN_PATH)
if not df_epic.empty:
    df_epic['OrderID_Clean'] = df_epic['Order ID'].astype(str).str.strip().str.replace(".0", "", regex=False)

df_flow = load_excel_safe(EPIC_FLOW_PATH)
if not df_flow.empty:
    df_flow['OrderID_Clean'] = df_flow['Order Med ID'].astype(str).str.strip().str.replace(".0", "", regex=False)

# Check Pyxis
p_sub = df_pyxis[df_pyxis['OrderNumber_Clean'] == order_id]
print(f"Pyxis transactions: {len(p_sub)}")
if not p_sub.empty:
    print(p_sub[['TransactionDateTime', 'TransactionType', 'DispenseAmount', 'WasteAmount']])

# Check Epic Administrations
e_sub = pd.DataFrame()
if not df_epic.empty:
    e_sub = df_epic[df_epic['OrderID_Clean'] == order_id]
    print(f"\nEpic Administrations: {len(e_sub)}")
    if not e_sub.empty:
        print(e_sub[['AdministrationInstant', 'AdministrationAction', 'Dose', 'DoseUnit']])

# Check Flowsheet
f_sub = pd.DataFrame()
if not df_flow.empty:
    f_sub = df_flow[df_flow['OrderID_Clean'] == order_id]
    print(f"\nEpic Flowsheets: {len(f_sub)}")
    if not f_sub.empty:
        print(f_sub[['Recorded Time', 'Flowsheet Row Name', 'Flowsheet Value']].head(10))

# Calculate elapsed duration
t_list = []
if not p_sub.empty:
    t_list.extend(pd.to_datetime(p_sub['TransactionDateTime']).dropna().tolist())
if not e_sub.empty:
    t_list.extend(pd.to_datetime(e_sub['AdministrationInstant']).dropna().tolist())
if not f_sub.empty:
    t_list.extend(pd.to_datetime(f_sub['Recorded Time']).dropna().tolist())

if t_list:
    min_t = min(t_list)
    max_t = max(t_list)
    elapsed = (max_t - min_t).total_seconds() / 3600.0
    print(f"\nTimestamps range: {min_t} to {max_t}")
    print(f"Elapsed duration (hours): {elapsed:.2f}")
else:
    print("\nNo timestamps found for this order!")
