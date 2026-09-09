import json

log_path = r"C:\Users\GaneshBhat\.gemini\antigravity-ide\brain\dd9f4209-9754-486e-8b4d-b09e26e62a72\.system_generated\logs\transcript.jsonl"

print("Searching transcript...")
with open(log_path, 'r', encoding='utf-8') as f:
    for line_num, line in enumerate(f, 1):
        try:
            data = json.loads(line)
            content = str(data.get("content", ""))
            tool_calls = str(data.get("tool_calls", ""))
            
            # Look for occurrences of recon_headers or format_sheet
            if "recon_headers" in content or "recon_headers" in tool_calls:
                print(f"\n--- Found on Line {line_num} ---")
                print("Step Index:", data.get("step_index"))
                print("Source:", data.get("source"))
                print("Type:", data.get("type"))
                
                # Check for replacements or code blocks
                if "def format_sheet" in content:
                    print("Contains format_sheet definition!")
                
                # Print a snippet of content
                print("Content preview (first 1000 chars):")
                print(content[:1000])
                print("-" * 50)
        except Exception as e:
            # Ignore bad lines
            pass
