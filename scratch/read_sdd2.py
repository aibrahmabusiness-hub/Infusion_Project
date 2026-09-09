import docx
import os

def read_docx(file_path):
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return
        
    doc = docx.Document(file_path)
    for i, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        if text:
            print(f"[{p.style.name}] {text}")

if __name__ == '__main__':
    read_docx(r"C:\Users\GaneshBhat\OneDrive - Novatio Solutions\Desktop\Infusion\Project Docs\SDD Infusion.docx")
