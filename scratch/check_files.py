import os

admin_path = r"c:\Users\GaneshBhat\OneDrive - Novatio Solutions\Desktop\Infusion\Input\BCH_Administered_Medications_Report_20260617_1354.xlsx"
flow_path = r"c:\Users\GaneshBhat\OneDrive - Novatio Solutions\Desktop\Infusion\Input\BCH_Flowsheet_Med_Infusion_Volume_20260617_1342.xlsx"
csv_path = r"c:\Users\GaneshBhat\OneDrive - Novatio Solutions\Desktop\Infusion\Input\AuditTransactionDetail_RC.csv"

print("Admin path exists:", os.path.exists(admin_path))
if os.path.exists(admin_path):
    print("Admin size:", os.path.getsize(admin_path))

print("Flow path exists:", os.path.exists(flow_path))
if os.path.exists(flow_path):
    print("Flow size:", os.path.getsize(flow_path))

print("CSV path exists:", os.path.exists(csv_path))
if os.path.exists(csv_path):
    print("CSV size:", os.path.getsize(csv_path))

print("Files in Input directory:")
for f in os.listdir(r"c:\Users\GaneshBhat\OneDrive - Novatio Solutions\Desktop\Infusion\Input"):
    print(" -", f)
