import os

for root, dirs, files in os.walk(r"c:\Users\GaneshBhat\OneDrive - Novatio Solutions\Desktop\Infusion"):
    for file in files:
        if file.endswith(".xlsx"):
            print(os.path.join(root, file))
