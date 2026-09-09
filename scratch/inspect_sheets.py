import openpyxl
import os
import shutil

def get_sheet_names_safe(path):
    temp_dir = os.path.dirname(path)
    temp_name = f"~temp_sheets_{os.path.basename(path)}"
    temp_path = os.path.join(temp_dir, temp_name)
    try:
        shutil.copy(path, temp_path)
    except Exception:
        import subprocess
        subprocess.run(['powershell', '-Command', f'Copy-Item -Path "{path}" -Destination "{temp_path}" -Force'], shell=True)
    wb = openpyxl.load_workbook(temp_path, read_only=True)
    sheets = wb.sheetnames
    try:
        os.remove(temp_path)
    except:
        pass
    return sheets

flow_path = r"c:\Users\GaneshBhat\OneDrive - Novatio Solutions\Desktop\Infusion\Input\BCH_Flowsheet_Med_Infusion_Volume_20260617_1342.xlsx"
print("Sheets in Flowsheet file:", get_sheet_names_safe(flow_path))
