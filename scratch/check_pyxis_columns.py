import pandas as pd

pyxis_path = r"c:\Users\GaneshBhat\OneDrive - Novatio Solutions\Desktop\Infusion\Input\AuditTransactionDetail_RC.csv"
df = pd.read_csv(pyxis_path)

print("=== columns ===")
for c in df.columns:
    # Print columns that have 'amt', 'amount', 'qty', 'quantity', 'waste', 'dispense'
    c_lower = c.lower()
    if any(x in c_lower for x in ['amt', 'amount', 'qty', 'quantity', 'waste', 'dispense']):
        print(f"Column: {c} | Sample: {df[c].dropna().head(3).tolist()}")

print("\n=== Sample of Waste rows ===")
waste_rows = df[df['TransactionType'].str.contains('Waste', na=False)]
print(waste_rows[['TransactionType', 'DispenseAmount', 'WasteAmount', 'Quantity', 'MedDescription']].head(5))

print("\n=== Sample of Vend rows ===")
vend_rows = df[df['TransactionType'].str.contains('Vend', na=False)]
print(vend_rows[['TransactionType', 'DispenseAmount', 'WasteAmount', 'Quantity', 'MedDescription']].head(5))
