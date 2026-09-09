import json

log_path = r"C:\Users\GaneshBhat\.gemini\antigravity-ide\brain\dd9f4209-9754-486e-8b4d-b09e26e62a72\.system_generated\logs\transcript.jsonl"

with open(log_path, 'r', encoding='utf-8') as f:
    for line_num, line in enumerate(f, 1):
        if line_num == 1485:
            data = json.loads(line)
            print("Content:")
            print(data.get("content"))
            break
