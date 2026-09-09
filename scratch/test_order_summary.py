import pandas as pd
import os

path = r'c:\Users\GaneshBhat\OneDrive - Novatio Solutions\Desktop\Infusion\Output\Step1_Pyxis_Analysis_7-13-2026_7-24-06_AM.xlsx'
if os.path.exists(path):
    xl = pd.ExcelFile(path)
    print("Sheets in Step1_Pyxis_Analysis_7-13-2026_7-24-06_AM.xlsx:")
    print(xl.sheet_names)
    
    df_recon = pd.read_excel(path, sheet_name="Order Reconciliation Summary")
    print("\nDoes 6078996223 exist in Order Reconciliation Summary?")
    print(df_recon[df_recon["Order ID"].astype(str).str.contains("6078996223")])
    
    df_raw = pd.read_excel(path, sheet_name="CS continuous infusions")
    print("\nDoes 6078996223 exist in CS continuous infusions?")
    p_sub = df_raw[df_raw["Order Number"].astype(str).str.contains("6078996223")]
    print(p_sub[["Order Number", "Transaction Type", "DispenseAmount (Raw)", "WasteAmount (Raw)", "Pairing Status"]])
else:
    print("Analysis file does not exist!")
