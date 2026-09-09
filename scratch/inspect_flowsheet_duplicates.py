import pandas as pd
import os
import shutil

def load_excel_safe(path):
    temp_dir = os.path.dirname(path)
    temp_name = f"~temp_{os.path.basename(path)}"
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

flow_path = r"c:\Users\GaneshBhat\OneDrive - Novatio Solutions\Desktop\Infusion\Input\BCH_Flowsheet_Med_Infusion_Volume_20260617_1342.xlsx"
df = load_excel_safe(flow_path)

print("Shape:", df.shape)
print("Columns:", df.columns.tolist())
print("Number of duplicate rows:", df.duplicated().sum())

# Check duplicates by CSN, Recorded Time, Flowsheet Row Name, Flowsheet Value
sub_dup = df.duplicated(subset=['CSN', 'Flowsheet Row Name', 'Recorded Time', 'Flowsheet Value']).sum()
print("Duplicates on key fields:", sub_dup)

print("\nSample rows:")
print(df.head(5))
