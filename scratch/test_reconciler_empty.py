import pandas as pd
import os

EPIC_ADMIN_PATH = r"c:\Users\GaneshBhat\OneDrive - Novatio Solutions\Desktop\Infusion\Input\BCH_Administered_Medications_Report_20260617_1354.xlsx"

# Replication of the reconciler.py code
if os.path.exists(EPIC_ADMIN_PATH):
    df_admin_raw = pd.read_excel(EPIC_ADMIN_PATH)
else:
    df_admin_raw = pd.DataFrame(columns=['Order ID', 'MRN', 'Patient', 'Medication', 'AdministrationAction', 'AdministrationInstant', 'Dose', 'DoseUnit', 'WeightAtRelease_X'])

print("Raw Columns:", df_admin_raw.columns.tolist())

def clean_id(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s

def clean_patient_name(name):
    return str(name).strip().upper()

def safe_to_datetime(val):
    return pd.to_datetime(val, errors='coerce')

def is_controlled(med):
    return True

df_admin = df_admin_raw.copy()
df_admin['Order ID_Clean'] = df_admin['Order ID'].apply(clean_id)
df_admin['MRN_Clean'] = df_admin['MRN'].apply(clean_id)
df_admin['Patient_Norm'] = df_admin['Patient'].apply(clean_patient_name)
df_admin['AdministrationInstant_Clean'] = df_admin['AdministrationInstant'].apply(safe_to_datetime)
print("Before filter Columns:", df_admin.columns.tolist())

df_admin = df_admin[df_admin['Medication'].apply(is_controlled)].copy()
print("After filter Columns:", df_admin.columns.tolist())
