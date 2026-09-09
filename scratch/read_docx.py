import docx
import sys

def read_docx(file_path):
    doc = docx.Document(file_path)
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip():
            print(f"Para {i}: {p.text}")

if __name__ == '__main__':
    read_docx(r"C:\Users\GaneshBhat\Downloads\Infusion Waste Monitoring - SDD.docx")
