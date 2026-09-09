import os
from datetime import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Color Palette (Corporate Slate/Navy Theme)
NAVY_HEADER_FILL = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
NAVY_HEADER_FONT = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")

SECTION_HEADER_FILL = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
SECTION_HEADER_FONT = Font(name="Segoe UI", size=12, bold=True, color="1A202C")

# Status Fills & Fonts (For highlighting rows/cells)
GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
GREEN_FONT = Font(name="Segoe UI", size=10, color="006100", bold=True)

RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
RED_FONT = Font(name="Segoe UI", size=10, color="9C0006", bold=True)

YELLOW_FILL = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
YELLOW_FONT = Font(name="Segoe UI", size=10, color="9C6500", bold=True)

ZEBRA_FILL = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

# Fonts
TITLE_FONT = Font(name="Segoe UI", size=16, bold=True, color="1B365D")
SUBTITLE_FONT = Font(name="Segoe UI", size=11, italic=True, color="4A5568")
LABEL_FONT = Font(name="Segoe UI", size=10, bold=True, color="4A5568")
DATA_FONT = Font(name="Segoe UI", size=10, color="000000")
BOLD_DATA_FONT = Font(name="Segoe UI", size=10, bold=True, color="000000")

# Borders
THIN_BORDER_SIDE = Side(border_style="thin", color="CBD5E1")
THIN_BORDER = Border(
    left=THIN_BORDER_SIDE,
    right=THIN_BORDER_SIDE,
    top=THIN_BORDER_SIDE,
    bottom=THIN_BORDER_SIDE
)

KPI_BORDER = Border(
    left=Side(border_style="medium", color="1B365D"),
    right=Side(border_style="medium", color="1B365D"),
    top=Side(border_style="medium", color="1B365D"),
    bottom=Side(border_style="medium", color="1B365D")
)

DOUBLE_BOTTOM_BORDER = Border(
    top=Side(border_style="thin", color="94A3B8"),
    bottom=Side(border_style="double", color="1B365D")
)

def format_sheet_common(ws, headers, df_data, title=None, subtitle=None, header_comments=None):
    """
    Applies standard formatting to a datasheet: headers on row 1, data starting on row 2.
    """
    start_row = 1
    
    # Write Headers to row 1
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=start_row, column=col_idx, value=header)
        cell.fill = NAVY_HEADER_FILL
        cell.font = NAVY_HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER
        
        if header_comments and header in header_comments:
            from openpyxl.comments import Comment
            comment = Comment(text=header_comments[header], author="System")
            comment.width = 250
            comment.height = 60
            cell.comment = comment
        
    ws.row_dimensions[start_row].height = 28
    
    # Write Data starting from row 2
    for row_idx, row_data in enumerate(df_data.values, start=start_row + 1):
        ws.row_dimensions[row_idx].height = 20
        # Determine zebra fill
        fill_to_use = ZEBRA_FILL if row_idx % 2 == 0 else WHITE_FILL
        
        for col_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = DATA_FONT
            cell.fill = fill_to_use
            cell.border = THIN_BORDER
            
            # Formatting and Alignment based on data type
            if isinstance(val, (int, float)):
                h_name = str(headers[col_idx - 1]).lower().strip()
                if "percentage" in h_name or "pct" in h_name or "%" in h_name:
                    cell.number_format = '0.0%'
                elif "duration" in h_name or "hours" in h_name:
                    cell.number_format = '#,##0.0'
                else:
                    cell.number_format = '#,##0.00' if isinstance(val, float) else '#,##0'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif hasattr(val, 'strftime'):  # datetime objects
                cell.value = val.strftime('%Y-%m-%d %H:%M')
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                # Text values
                cell.alignment = Alignment(horizontal="left", vertical="center")
                # Format IDs and status as centered
                val_str = str(val).strip()
                if val_str.isdigit() and len(val_str) > 5:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.number_format = '@'
                elif val_str in ['Reconciled', 'Discrepancy', 'Exception', 'Matched', 'Unmatched']:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    
    # 4. Auto-fit column widths
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        
        # Don't size based on title rows
        cells_to_check = col[start_row-1:] if start_row > 1 else col
        for cell in cells_to_check:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

def build_summary_dashboard(ws, df_reconcile, df_exceptions):
    """
    Creates a styled summary dashboard tab with KPI blocks and drug-level overview.
    """
    ws.title = "Summary Dashboard"
    ws.views.sheetView[0].showGridLines = True
    
    # Title Block
    ws.cell(row=2, column=2, value="Controlled Substance Infusion Reconciliation Report").font = TITLE_FONT
    ws.cell(row=3, column=2, value=f"Generated on {datetime.now().strftime('%Y-%m-%d %H:%M')} | Pharmacy Auditing Dashboard").font = SUBTITLE_FONT
    
    # Calculate KPIs
    total_orders = len(df_reconcile)
    reconciled_orders = len(df_reconcile[df_reconcile['OrderStatus'] == 'Reconciled'])
    discrepant_orders = len(df_reconcile[df_reconcile['OrderStatus'] == 'Discrepancy'])
    exception_orders = len(df_reconcile[df_reconcile['OrderStatus'] == 'Exception'])
    
    # KPI 1: Reconciled
    ws.merge_cells("B5:D6")
    kpi1_val = ws.cell(row=5, column=2, value=reconciled_orders)
    kpi1_val.font = Font(name="Segoe UI", size=22, bold=True, color="006100")
    kpi1_val.alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=7, column=2, value="RECONCILED ORDERS (100% MATCH)").font = LABEL_FONT
    ws.cell(row=7, column=2).alignment = Alignment(horizontal="center")
    
    # KPI 2: Discrepant
    ws.merge_cells("F5:H6")
    kpi2_val = ws.cell(row=5, column=6, value=discrepant_orders)
    kpi2_val.font = Font(name="Segoe UI", size=22, bold=True, color="9C0006")
    kpi2_val.alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=7, column=6, value="ORDERS WITH MATERIAL VARIANCE").font = LABEL_FONT
    ws.cell(row=7, column=6).alignment = Alignment(horizontal="center")
    
    # KPI 3: Exceptions
    ws.merge_cells("J5:L6")
    kpi3_val = ws.cell(row=5, column=10, value=exception_orders)
    kpi3_val.font = Font(name="Segoe UI", size=22, bold=True, color="9C6500")
    kpi3_val.alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=7, column=10, value="EXCEPTIONS FLAGGED FOR REVIEW").font = LABEL_FONT
    ws.cell(row=7, column=10).alignment = Alignment(horizontal="center")
    
    # Style KPI boxes
    for start_col, end_col, color_fill in [(2, 4, GREEN_FILL), (6, 8, RED_FILL), (10, 12, YELLOW_FILL)]:
        for r in range(5, 7):
            for c in range(start_col, end_col + 1):
                ws.cell(row=r, column=c).fill = color_fill
                ws.cell(row=r, column=c).border = KPI_BORDER
                
    # Drug Level Overview Table
    ws.cell(row=9, column=2, value="Medication Summary Table").font = SECTION_HEADER_FONT
    ws.cell(row=9, column=2).fill = SECTION_HEADER_FILL
    ws.merge_cells("B9:H9")
    ws.row_dimensions[9].height = 24
    
    dashboard_headers = [
        "Medication", "Total Orders", "Reconciled", "Discrepancies", "Exceptions", 
        "Dispensed Amount", "Administered Amount", "Variance"
    ]
    for col_idx, h in enumerate(dashboard_headers, start=2):
        cell = ws.cell(row=10, column=col_idx, value=h)
        cell.fill = NAVY_HEADER_FILL
        cell.font = NAVY_HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER
    ws.row_dimensions[10].height = 24
    
    # Compute Medication Level Summary Data
    summary_data = []
    # Clean Medication name to group by base drug name
    df_reconcile_cp = df_reconcile.copy()
    df_reconcile_cp['BaseDrug'] = df_reconcile_cp['Medication'].apply(lambda x: str(x).split(' ')[0].split('(')[0].strip().upper())
    
    for drug_name, group in df_reconcile_cp.groupby('BaseDrug'):
        if not drug_name or drug_name == "UNKNOWN" or pd.isna(drug_name):
            continue
            
        cnt = len(group)
        rec = len(group[group['OrderStatus'] == 'Reconciled'])
        disc = len(group[group['OrderStatus'] == 'Discrepancy'])
        ex = len(group[group['OrderStatus'] == 'Exception'])
        
        # Get primary reconciliation unit for unit sum display
        rep_unit = group.iloc[0]['ReconciliationUnit']
        
        # Calculate totals in reconciliation unit
        tot_disp = 0.0
        tot_admin = 0.0
        tot_var = 0.0
        for _, row in group.iterrows():
            tot_disp += row['TotalDispensed']
            tot_admin += row['TotalAdministered']
            tot_var += row['AdditionalInWaste']
                
        summary_data.append({
            'Medication': f"{drug_name} ({rep_unit})",
            'TotalOrders': cnt,
            'Reconciled': rec,
            'Discrepancies': disc,
            'Exceptions': ex,
            'Dispensed': tot_disp,
            'Administered': tot_admin,
            'Variance': tot_var
        })
        
    # Write summary rows
    current_row = 11
    for row_idx, row_dict in enumerate(summary_data, start=11):
        ws.row_dimensions[row_idx].height = 20
        fill_to_use = ZEBRA_FILL if row_idx % 2 == 0 else WHITE_FILL
        
        ws.cell(row=row_idx, column=2, value=row_dict['Medication']).font = BOLD_DATA_FONT
        ws.cell(row=row_idx, column=3, value=row_dict['TotalOrders']).alignment = Alignment(horizontal="right")
        ws.cell(row=row_idx, column=4, value=row_dict['Reconciled']).alignment = Alignment(horizontal="right")
        ws.cell(row=row_idx, column=5, value=row_dict['Discrepancies']).alignment = Alignment(horizontal="right")
        ws.cell(row=row_idx, column=6, value=row_dict['Exceptions']).alignment = Alignment(horizontal="right")
        ws.cell(row=row_idx, column=7, value=row_dict['Dispensed']).alignment = Alignment(horizontal="right")
        ws.cell(row=row_idx, column=8, value=row_dict['Administered']).alignment = Alignment(horizontal="right")
        ws.cell(row=row_idx, column=9, value=row_dict['Variance']).alignment = Alignment(horizontal="right")
        
        # Formats and borders
        for c in range(2, 10):
            cell = ws.cell(row=row_idx, column=c)
            cell.font = DATA_FONT if c != 2 else BOLD_DATA_FONT
            cell.fill = fill_to_use
            cell.border = THIN_BORDER
            if c > 2 and c <= 6:
                cell.number_format = '#,##0'
            elif c >= 7:
                cell.number_format = '#,##0.00'
                
            # If variance cell is non-zero, highlight it
            if c == 9 and abs(row_dict['Variance']) > 1e-5:
                cell.fill = RED_FILL
                cell.font = RED_FONT
                
        current_row += 1
        
    # Totals Row
    ws.row_dimensions[current_row].height = 22
    tot_label = ws.cell(row=current_row, column=2, value="TOTALS")
    tot_label.font = BOLD_DATA_FONT
    tot_label.alignment = Alignment(horizontal="left")
    tot_label.border = DOUBLE_BOTTOM_BORDER
    
    for c in range(3, 10):
        col_letter = get_column_letter(c)
        cell = ws.cell(row=current_row, column=c)
        if c in [3, 4, 5, 6]:
            cell.value = f"=SUM({col_letter}11:{col_letter}{current_row-1})"
            cell.number_format = '#,##0'
        elif c in [7, 8, 9]:
            # SUM of mixed units is medically not correct, but shows spreadsheet counts
            cell.value = f"=SUM({col_letter}11:{col_letter}{current_row-1})"
            cell.number_format = '#,##0.00'
            
        cell.font = BOLD_DATA_FONT
        cell.alignment = Alignment(horizontal="right")
        cell.border = DOUBLE_BOTTOM_BORDER
        
    # Set Dashboard column widths
    ws.column_dimensions['A'].width = 3
    ws.column_dimensions['B'].width = 30
    for c in ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
        ws.column_dimensions[c].width = 18

def generate_excel_report(data_dict, file_path):
    """
    Generates a beautifully styled and fully formatted Excel workbook.
    `data_dict` keys: 'reconciliation', 'pyxis_details', 'epic_admin_details', 'epic_flowsheet_details', 'exceptions'
    """
    wb = openpyxl.Workbook()
    
    # 1. Summary Dashboard Sheet
    ws_dash = wb.active
    build_summary_dashboard(ws_dash, data_dict['reconciliation'], data_dict['exceptions'])
    
    # 2. Reconciliation Summary Sheet
    ws_reconcile = wb.create_sheet(title="Order Reconciliation Summary")
    ws_reconcile.views.sheetView[0].showGridLines = True
    reconcile_headers = [
        "Order ID", "Patient Name", "Medication Name", "Reconciliation Unit",
        "Cumulative Dispensed", "Cumulative Initial Waste", "Available to be Administered",
        "Cumulative Administered", "Additional in Waste", "Order Status",
        "Has Unmatched Pyxis Transactions", "Cumulative Administered Status", "Administration Action"
    ]
    reconcile_header_comments = {
        "Available to be Administered": "Available to be Administered = Cumulative Dispensed (E) - Cumulative Initial Waste (F)",
        "Additional in Waste": "Additional in Waste = Available to be Administered (G) - Cumulative Administered (H)"
    }
    format_sheet_common(
        ws_reconcile, reconcile_headers, data_dict['reconciliation'],
        title="Medication Infusion Reconciliation Summary",
        subtitle="Order-level reconciliation summary tracking cumulative dispensed, administered, wasted, and variances",
        header_comments=reconcile_header_comments
    )
    
    # Apply conditional colors to Reconciliation Status and highlight rows
    for r in range(2, ws_reconcile.max_row + 1):
        status_val = str(ws_reconcile.cell(row=r, column=10).value).lower()
        # Highlights based on status
        if status_val == "reconciled":
            ws_reconcile.cell(row=r, column=10).fill = GREEN_FILL
            ws_reconcile.cell(row=r, column=10).font = GREEN_FONT
        elif status_val == "discrepancy":
            ws_reconcile.cell(row=r, column=10).fill = RED_FILL
            ws_reconcile.cell(row=r, column=10).font = RED_FONT
            # Highlight variance (Additional in Waste) column in light red too
            ws_reconcile.cell(row=r, column=9).fill = RED_FILL
            ws_reconcile.cell(row=r, column=9).font = RED_FONT
        elif status_val == "exception":
            ws_reconcile.cell(row=r, column=10).fill = YELLOW_FILL
            ws_reconcile.cell(row=r, column=10).font = YELLOW_FONT
            
    # 3. Administered Calculation Details Sheet
    ws_admin_calc = wb.create_sheet(title="Administered Calculation")
    admin_calc_headers = [
        "Order ID", "Patient Name", "Medication Name", "Reconciliation Unit",
        "Flowsheet Entries Count", "Total Administered Volume (mL)", "Concentration (per mL)",
        "Cumulative Administered Dose", "Final Cumulative Administered (Rec Unit)"
    ]
    admin_calc_header_comments = {
        "Total Administered Volume (mL)": "Sum of all matched Flowsheet recorded volume entries for this order.",
        "Cumulative Administered Dose": "Total Administered Volume (mL) * Concentration (per mL) (when Reconciliation Unit is in mg/mcg)",
        "Final Cumulative Administered (Rec Unit)": "Equal to Volume (mL) if unit is mL, otherwise Cumulative Administered Dose if unit is mg/mcg."
    }
    format_sheet_common(
        ws_admin_calc, admin_calc_headers, data_dict['administered_details'],
        title="Cumulative Administered Calculation Details",
        subtitle="Verification detail showing how the final cumulative administered amounts were calculated",
        header_comments=admin_calc_header_comments
    )

    # 4. Pyxis Details Sheet
    ws_pyxis = wb.create_sheet(title="Pyxis Transactions Log")
    ws_pyxis.views.sheetView[0].showGridLines = True
    pyxis_headers = [
        "Order Number", "Patient ID (MRN)", "Patient Name", "Transaction DateTime", "Transaction Type",
        "Transaction Status", "MedDescription", "DispenseAmount (Raw)", "WasteAmount (Raw)", "Parsed Conc/mL", "Parsed Unit",
        "Parsed Total Vol", "Calculated Dose (mg/mcg)", "Converted Vol (mL)", "Paired Time",
        "Paired Type", "Paired Amount", "Pairing Status", "Pairing Gap (Hours)"
    ]
    format_sheet_common(
        ws_pyxis, pyxis_headers, data_dict['pyxis_details'],
        title="Pyxis Transaction Audit Details",
        subtitle="Granular log of all Pyxis Vends and Wastes including regex parsing results and closest timestamp pairings"
    )
    # Highlight pairing statuses
    for r in range(2, ws_pyxis.max_row + 1):
        pair_status = ws_pyxis.cell(row=r, column=18).value
        if pair_status == "Matched":
            ws_pyxis.cell(row=r, column=18).fill = GREEN_FILL
            ws_pyxis.cell(row=r, column=18).font = GREEN_FONT
        elif pair_status == "Unmatched - Missing Waste":
            ws_pyxis.cell(row=r, column=18).fill = RED_FILL
            ws_pyxis.cell(row=r, column=18).font = RED_FONT
        elif pair_status == "Unmatched - Excess Waste":
            ws_pyxis.cell(row=r, column=18).fill = YELLOW_FILL
            ws_pyxis.cell(row=r, column=18).font = YELLOW_FONT
            
    # 4. Epic Admin Log Sheet
    ws_admin = wb.create_sheet(title="Epic Administrations Log")
    ws_admin.views.sheetView[0].showGridLines = True
    admin_headers = [
        "Order ID", "MRN", "Patient Name", "Medication", "Administrative Action", "Administration Instant",
        "Dose", "Dose Unit", "Patient Weight (kg)"
    ]
    format_sheet_common(
        ws_admin, admin_headers, data_dict['epic_admin_details'],
        title="Epic Medication Administrations Log",
        subtitle="Clinical record of administration actions (New Bag, Given, Bolus, Stopped, Restarted) providing context"
    )
    # Highlight Stopped and Restarted rows
    for r in range(2, ws_admin.max_row + 1):
        action_val = str(ws_admin.cell(row=r, column=5).value).lower()
        if 'stop' in action_val:
            ws_admin.cell(row=r, column=5).fill = RED_FILL
            ws_admin.cell(row=r, column=5).font = RED_FONT
        elif 'restart' in action_val:
            ws_admin.cell(row=r, column=5).fill = YELLOW_FILL
            ws_admin.cell(row=r, column=5).font = YELLOW_FONT
            
    # 5. Epic Flowsheet Log Sheet
    ws_flow = wb.create_sheet(title="Epic Flowsheet Hourly Log")
    ws_flow.views.sheetView[0].showGridLines = True
    flow_headers = [
        "CSN", "MRN", "Patient Name", "Order ID", "Flowsheet Row Name", "Display Name",
        "Recorded Time", "Flowsheet Value (mL)"
    ]
    format_sheet_common(
        ws_flow, flow_headers, data_dict['epic_flowsheet_details'],
        title="Epic Flowsheet Hourly Infusion Volumes",
        subtitle="Hourly logs of infusion volumes (mL) recorded by clinicians in the Patient Flowsheet"
    )
    
    # 6. Exceptions Log Sheet
    ws_exceptions = wb.create_sheet(title="Audit Exceptions Log")
    ws_exceptions.views.sheetView[0].showGridLines = True
    exceptions_headers = [
        "Order ID", "MRN", "Patient Name", "Source File", "Exception Type", "Record Identifier",
        "Detailed Reason", "Pharmacy Review Action Required"
    ]
    format_sheet_common(
        ws_exceptions, exceptions_headers, data_dict['exceptions'],
        title="Pharmacy Audit Exception Log",
        subtitle="Consolidated lists of clinical discrepancies, missing documentation, workflow gaps, and diversion risks"
    )
    # Highlight Exception Rows in light yellow
    for r in range(2, ws_exceptions.max_row + 1):
        for c in range(1, 9):
            ws_exceptions.cell(row=r, column=c).fill = YELLOW_FILL
        ws_exceptions.cell(row=r, column=5).font = YELLOW_FONT
        
    # Save file
    print(f"Saving workbook to {file_path}...")
    wb.save(file_path)
    print("Report generated successfully!")
