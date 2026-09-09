import docx
import os

def update_document(file_path, output_path):
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return

    doc = docx.Document(file_path)

    for p in doc.paragraphs:
        # Update Source 1: Pyxis to include Controlled Substances
        if "Key Fields" in p.text and "Patient Name (Column E)" in [para.text for para in doc.paragraphs]: # Just finding a safe place, but actually let's just do exact string replacement.
            pass
            
        # Update Reconciliation Logic
        if "Total Dispensed Amount =" in p.text:
            p.text = "Total Dispensed Amount =\nSUM(Paired Vend Amounts)"
        elif "SUM(All Vend Amounts)" in p.text:
            p.text = p.text.replace("SUM(All Vend Amounts)", "SUM(Paired Vend Amounts)")
            
        if "Total Waste Amount =" in p.text:
            p.text = "Total Waste Amount =\nSUM(Paired Waste Amounts)"
        elif "SUM(All Waste Amounts)" in p.text:
            p.text = p.text.replace("SUM(All Waste Amounts)", "SUM(Paired Waste Amounts)")

    # Adding a new paragraph about Vend/Waste Pairing before Date Matching Logic
    # We will find "Date Matching Logic" and insert before it
    for i, p in enumerate(doc.paragraphs):
        if p.text == "Date Matching Logic":
            new_p = p.insert_paragraph_before("Vend/Waste Pairing Logic")
            new_p.style = p.style
            
            new_p2 = p.insert_paragraph_before("Unlike simple summation, Vends and Wastes must be chronologically paired. A Waste transaction is paired with a preceding Vend transaction that occurred within the maximum hang time window (e.g. 96 hours).")
            new_p2.style = doc.styles['Normal']
            
            new_p3 = p.insert_paragraph_before("Transaction types 'Stock Return' and 'Vend (Cancel)' are excluded from active pairs. Any unmatched Wastes are flagged as exceptions.")
            new_p3.style = doc.styles['Normal']
            break
            
    # Find "Source File 1: Pyxis" or similar to add controlled substance logic
    for i, p in enumerate(doc.paragraphs):
        if p.text == "Key Fields" and i > 0 and "Pyxis" in doc.paragraphs[i-1].text:
            pass # Just an anchor
            
    # Just add controlled substance note under Pyxis section
    for i, p in enumerate(doc.paragraphs):
        if "The documented waste amount is found in Quantity" in p.text:
            new_p = p.insert_paragraph_before("Controlled Substances Filtering:")
            new_p.style = p.style
            new_p2 = p.insert_paragraph_before("The system exclusively filters for controlled substances (e.g., fentanyl, midazolam, versed, ketamine, morphine, hydromorphone, dilaudid) using the MedDescription field. Non-controlled substances are excluded from reconciliation.")
            new_p2.style = doc.styles['Normal']
            break
            
    doc.save(output_path)
    print(f"Updated document saved to {output_path}")

if __name__ == '__main__':
    input_doc = r"C:\Users\GaneshBhat\OneDrive - Novatio Solutions\Desktop\Infusion\Project Docs\SDD Infusion.docx"
    output_doc = r"C:\Users\GaneshBhat\OneDrive - Novatio Solutions\Desktop\Infusion\Project Docs\SDD Infusion_Updated.docx"
    update_document(input_doc, output_doc)
