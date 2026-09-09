import os
from datetime import datetime

conversations_dir = r"C:\Users\GaneshBhat\.gemini\antigravity-ide\conversations"
files = []
for name in os.listdir(conversations_dir):
    path = os.path.join(conversations_dir, name)
    mtime = os.path.getmtime(path)
    files.append((name, mtime, os.path.getsize(path)))

files.sort(key=lambda x: x[1], reverse=True)

print("=== RECENT CONVERSATIONS ===")
for name, mtime, size in files[:10]:
    print(f"File: {name} | Size: {size} bytes | Modified: {datetime.fromtimestamp(mtime).isoformat()}")
