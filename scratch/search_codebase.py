import os
import re

query = 'cumulative'
root_dir = r"c:\Users\GaneshBhat\OneDrive - Novatio Solutions\Desktop\Infusion"

for dirpath, dirnames, filenames in os.walk(root_dir):
    if '.git' in dirpath or '__pycache__' in dirpath or '.gemini' in dirpath:
        continue
    for f in filenames:
        if f.endswith('.py'):
            path = os.path.join(dirpath, f)
            try:
                with open(path, 'r', encoding='utf-8', errors='ignore') as file:
                    for line_num, line in enumerate(file, 1):
                        if query.lower() in line.lower():
                            print(f"{os.path.relpath(path, root_dir)}:{line_num}: {line.strip()}")
            except Exception as e:
                print(f"Error reading {path}: {e}")
