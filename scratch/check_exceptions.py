import os
import sys
# Dynamically add the root directory to path
sys.path.append(r"c:\Users\GaneshBhat\OneDrive - Novatio Solutions\Desktop\Infusion")

import pandas as pd
from src.reconciler import run_reconciliation

results = run_reconciliation()
df_rec = results['reconciliation']
df_exc = results['exceptions']

print("=== RECONCILIATION STATUS COUNTS ===")
print(df_rec['ReconciliationStatus'].value_counts())

print("\n=== EXCEPTION TYPE COUNTS ===")
if not df_exc.empty:
    print(df_exc['ExceptionType'].value_counts())
    
    print("\n=== SAMPLE EXCEPTIONS (FIRST 10) ===")
    for idx, row in df_exc.head(10).iterrows():
        print(f"OrderID: {row['OrderID']} | Source: {row['SourceFile']} | Type: {row['ExceptionType']}")
        print(f"  Description: {row['Description']}")
        print(f"  ActionNeeded: {row['ActionNeeded']}")
        print("-" * 50)
else:
    print("No exceptions found.")
