import os

base_dir = r"c:\Users\GaneshBhat\OneDrive - Novatio Solutions\Desktop\Infusion"
print("=== Searching for n8n/webhook references in codebase ===")
for root, dirs, files in os.walk(base_dir):
    for file in files:
        if file.endswith((".py", ".ps1", ".bat", ".json", ".txt")):
            path = os.path.join(root, file)
            try:
                with open(path, "r", encoding="utf-8", errors="ignore") as f:
                    content = f.read()
                    if "n8n" in content.lower() or "webhook" in content.lower():
                        rel = os.path.relpath(path, base_dir)
                        print(f"Found in: {rel}")
            except Exception as e:
                pass
