import os
import pandas as pd

input_dir = r"c:\Users\GaneshBhat\OneDrive - Novatio Solutions\Desktop\Infusion\Input"

for filename in os.listdir(input_dir):
    filepath = os.path.join(input_dir, filename)
    if filename.endswith('.xlsx'):
        try:
            xl = pd.ExcelFile(filepath)
            print(f"File: {filename}")
            print(f"  Sheets: {xl.sheet_names}")
            for sheet in xl.sheet_names:
                df = pd.read_excel(filepath, sheet_name=sheet, nrows=5)
                print(f"    Sheet '{sheet}' shape: {df.shape}")
                print(f"    Columns: {list(df.columns)}")
        except Exception as e:
            print(f"Error reading {filename}: {e}")
    elif filename.endswith('.csv'):
        try:
            df = pd.read_csv(filepath, nrows=5)
            print(f"File: {filename}")
            print(f"  Columns: {list(df.columns)}")
        except Exception as e:
            print(f"Error reading {filename}: {e}")
