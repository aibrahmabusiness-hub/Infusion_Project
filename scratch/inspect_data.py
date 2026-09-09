import pandas as pd

pyxis_path = r"c:\Users\GaneshBhat\OneDrive - Novatio Solutions\Desktop\Infusion\Input\AuditTransactionDetail_RC.csv"
admin_path = r"c:\Users\GaneshBhat\OneDrive - Novatio Solutions\Desktop\Infusion\Input\BCH_Administered_Medications_Report_20260617_1354.xlsx"
flow_path = r"c:\Users\GaneshBhat\OneDrive - Novatio Solutions\Desktop\Infusion\Input\BCH_Flowsheet_Med_Infusion_Volume_20260617_1342.xlsx"

df_pyxis = pd.read_csv(pyxis_path)
df_admin = pd.read_excel(admin_path)
df_flow = pd.read_excel(flow_path)

print("=== Raw Columns ===")
print("Pyxis:", df_pyxis.columns.tolist()[:10])
print("Epic Admin:", df_admin.columns.tolist()[:10])
print("Flowsheet:", df_flow.columns.tolist()[:10])

print("\n=== Shapes ===")
print("Pyxis:", df_pyxis.shape)
print("Epic Admin:", df_admin.shape)
print("Flowsheet:", df_flow.shape)

print("\n=== Unique Order IDs in Raw Data ===")
pyxis_orders = df_pyxis['OrderNumber'].dropna().unique()
admin_orders = df_admin['Order ID'].dropna().unique()
flow_orders = df_flow['Order Med ID'].dropna().unique()

print(f"Pyxis Unique Order Numbers: {len(pyxis_orders)}")
print(f"Epic Admin Unique Order IDs: {len(admin_orders)}")
print(f"Flowsheet Unique Order Med IDs: {len(flow_orders)}")

# Check clean ID function output
def clean_id(val):
    import numpy as np
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s

pyxis_clean = set(clean_id(x) for x in pyxis_orders) - {''}
admin_clean = set(clean_id(x) for x in admin_orders) - {''}
flow_clean = set(clean_id(x) for x in flow_orders) - {''}

print("\n=== Cleaned Unique Order IDs ===")
print(f"Pyxis: {len(pyxis_clean)}")
print(f"Epic Admin: {len(admin_clean)}")
print(f"Flowsheet: {len(flow_clean)}")

print("\n=== Overlap between Pyxis and Flowsheet ===")
print(f"Intersection of Pyxis and Flowsheet: {len(pyxis_clean.intersection(flow_clean))}")
print(f"In Flowsheet but NOT in Pyxis: {len(flow_clean - pyxis_clean)}")
print(f"In Pyxis but NOT in Flowsheet: {len(pyxis_clean - flow_clean)}")

print("\n=== Sample Order IDs in Pyxis ===")
print(list(pyxis_clean)[:10])
print("\n=== Sample Order IDs in Flowsheet ===")
print(list(flow_clean)[:10])
print("\n=== Sample Order IDs in Admin ===")
print(list(admin_clean)[:10])
