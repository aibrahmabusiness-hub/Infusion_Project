import docx
import os

def update_sdd(input_file, output_file):
    doc = docx.Document(input_file)
    
    # Text replacements to reflect the latest logic
    replacements = {
        "An email will be received in the Intelligent Automation inbox with the subject: (TBD). It will contain three different email attachments:":
            "An email will be received in the Intelligent Automation inbox with the subject: (TBD). It will contain these email attachments, and the bot will load the following inputs:",
            
        "Flow Sheet": 
            "1. Flow Sheet (Epic Flowsheet)",
            
        "Medication Administrations": 
            "2. Medication Administrations (Epic Administrations - EPIC_ADMIN_PATH)",
            
        "All Device Event Report": 
            "3. All Device Event Report (Pyxis)",
            
        "Open the All Device Event Report_Pyxis.xlsx file.":
            "Open the Pyxis, Epic Administrations, and Epic Flowsheet files. A Historical State file is also loaded to track active carry-over orders.",
            
        "Identify and calculate the total waste for each order. (Add filter in column “TransactionType”)":
            "Filter Pyxis transactions to strictly Controlled Substances (e.g. fentanyl, midazolam). Group and Pair 'Vend' and 'Waste' transactions by Order ID and MedDescription, factoring in maximum hang time gaps. Exclude 'Stock Return' and 'Vend (Cancel)' from active pairs.",
            
        "(TBA) : Open Medication Administration file and identify bag is : new Bag or Bolus":
            "Process Epic Administration and Flowsheet records. Align the calculated Pyxis Administered Volume (Vend - paired Wastes) against Flowsheet infused totals and Epic Administration records.",
            
        "Match the unique Order ID and retrieve the corresponding total values from the flowsheet.\n(Order values should be aggregated, and rate logic is TBD.)":
            "Match the unique Order ID across all datasets to retrieve the aggregated totals from Flowsheet and Epic Administrations.",
            
        "Final waste= Remaining Medication  – Flowsheet Total Value (per order)":
            "Final Variance = Pyxis Administered Volume - Flowsheet Total Value (per order)"
    }
    
    # We will iterate through the paragraphs and apply replacements
    for p in doc.paragraphs:
        for old_text, new_text in replacements.items():
            if old_text in p.text:
                # To maintain some style, we could clear runs and add one new run, but it loses inline formatting.
                # However, for plain text bullet points, this is acceptable.
                p.text = p.text.replace(old_text, new_text)

    doc.save(output_file)
    print(f"Successfully updated document and saved to {output_file}")

if __name__ == '__main__':
    input_doc = r"C:\Users\GaneshBhat\Downloads\Infusion Waste Monitoring - SDD.docx"
    output_doc = r"C:\Users\GaneshBhat\Downloads\Infusion Waste Monitoring - SDD_Updated.docx"
    if os.path.exists(input_doc):
        update_sdd(input_doc, output_doc)
    else:
        print(f"Error: {input_doc} not found.")
