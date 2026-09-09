import pandas as pd
import numpy as np

pyxis_path = r"c:\Users\GaneshBhat\OneDrive - Novatio Solutions\Desktop\Infusion\Input\AuditTransactionDetail_RC.csv"
admin_path = r"c:\Users\GaneshBhat\OneDrive - Novatio Solutions\Desktop\Infusion\Input\BCH_Administered_Medications_Report_20260617_1354.xlsx"
flow_path = r"c:\Users\GaneshBhat\OneDrive - Novatio Solutions\Desktop\Infusion\Input\BCH_Flowsheet_Med_Infusion_Volume_20260617_1342.xlsx"

df_pyxis_raw = pd.read_csv(pyxis_path)
df_admin_raw = pd.read_excel(admin_path)
df_flow_raw = pd.read_excel(flow_path)

CONTROLLED_DRUGS = ['fentanyl', 'midazolam', 'versed', 'ketamine', 'morphine', 'hydromorphone', 'dilaudid']

def is_controlled(text):
    if not isinstance(text, str):
        return False
    text_lower = text.lower()
    return any(drug in text_lower for drug in CONTROLLED_DRUGS)

def clean_id(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s

df_pyxis = df_pyxis_raw.copy()
df_pyxis['OrderNumber_Clean'] = df_pyxis['OrderNumber'].apply(clean_id)
df_pyxis = df_pyxis[df_pyxis['MedDescription'].apply(is_controlled)].copy()
df_pyxis['TransactionType_Clean'] = df_pyxis['TransactionType'].astype(str).str.strip()
df_pyxis = df_pyxis[df_pyxis['TransactionType_Clean'].isin(['Vend', 'Waste', 'Stock Return', 'Vend (Cancel)'])].copy()

df_admin = df_admin_raw.copy()
df_admin['Order ID_Clean'] = df_admin['Order ID'].apply(clean_id)
df_admin = df_admin[df_admin['Medication'].apply(is_controlled)].copy()

df_flow = df_flow_raw.copy()
df_flow['Order Med ID_Clean'] = df_flow['Order Med ID'].apply(clean_id)
df_flow = df_flow[df_flow['Flowsheet Row Name'].apply(is_controlled)].copy()

pyxis_orders = set(df_pyxis['OrderNumber_Clean'].dropna().unique()) - {''}
admin_orders = set(df_admin['Order ID_Clean'].dropna().unique()) - {''}
flow_orders = set(df_flow['Order Med ID_Clean'].dropna().unique()) - {''}
all_orders = pyxis_orders.union(flow_orders)

print(f"Filtered Pyxis unique orders: {len(pyxis_orders)}")
print(f"Filtered Admin unique orders: {len(admin_orders)}")
print(f"Filtered Flowsheet unique orders: {len(flow_orders)}")
print(f"Filtered Union (all_orders): {len(all_orders)}")

# Check if there are any non-controlled records in Pyxis, Admin, Flowsheet that are related
print("\nIs it possible that flowsheet contains controlled drugs but they are named differently?")
print("Example flowsheet rows:")
print(df_flow_raw['Flowsheet Row Name'].dropna().unique()[:30])

print("\nLet's check if there are matches in flow_orders that are in admin_orders but not in pyxis_orders:")
flow_in_admin = flow_orders.intersection(admin_orders)
print(f"Flow orders in Admin: {len(flow_in_admin)}")
flow_in_pyxis = flow_orders.intersection(pyxis_orders)
print(f"Flow orders in Pyxis: {len(flow_in_pyxis)}")
print(f"Admin orders in Pyxis: {len(admin_orders.intersection(pyxis_orders))}")
