import pandas as pd
from tkinter import filedialog, Tk
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# Hide tkinter window
Tk().withdraw()


export_file = filedialog.askopenfilename(title="Select Export Excel File (sheet name 'Table')")
if not export_file:
    print("❌ Export file not selected.")
    exit()

# Select Mapping File
mapping_file = filedialog.askopenfilename(title="Select Target Mapping File (sheet name 'MAPPING')")
if not mapping_file:
    print("❌ Mapping file not selected.")
    exit()

try:
    #  Read sheets
    export_df = pd.read_excel(export_file, sheet_name="Table")
    mapping_df = pd.read_excel(mapping_file, sheet_name="MAPPING")

  
    export_df['deposit_date'] = pd.to_datetime(export_df['deposit_date'], errors='coerce')
    export_df = export_df.dropna(subset=['deposit_date'])
    export_df['only_date'] = export_df['deposit_date'].dt.date

    export_df['type'] = export_df['type'].astype(str).str.strip().str.lower()
    export_df['status'] = export_df['status'].astype(str).str.strip().str.lower()

  
    latest_date = export_df['only_date'].max()
    print("🗓️ Latest Date:", latest_date)

  
    filtered_export = export_df[
        (export_df['only_date'] == latest_date) &
        (export_df['type'].isin(['card', 'cash'])) &
        (export_df['status'] == 'accepted')
    ]

    #  Merge with mapping for latest amount data
    merged = pd.merge(filtered_export, mapping_df, on='code', how='inner')

    
    mapping_df = mapping_df[mapping_df['target'].fillna(0) != 0]

   
    grouped_amount = (
        merged.groupby('asmname', as_index=False)['amount'].sum()
        .rename(columns={'asmname': 'ASM Name', 'amount': 'Amount'})
    )

    
    unique_target = (
        mapping_df.drop_duplicates(subset='asmname')[['asmname', 'target']]
        .rename(columns={'asmname': 'ASM Name', 'target': 'Target'})
    )

    
    final = pd.merge(unique_target, grouped_amount, on='ASM Name', how='left')
    final['Amount'] = final['Amount'].fillna(0)


    final['% Achievement'] = final.apply(
        lambda row: (row['Amount'] / row['Target']) if row['Target'] != 0 else 0, axis=1
    )

    
    total_target = final['Target'].sum()
    total_amount = final['Amount'].sum()
    total_achieved = (total_amount / total_target) if total_target != 0 else 0

    total_row = {
        'ASM Name': 'TOTAL',
        'Target': total_target,
        'Amount': total_amount,
        '% Achievement': total_achieved
    }

    final.loc[len(final)] = total_row

    
    data_only = final[final['ASM Name'] != 'TOTAL'].sort_values(by='% Achievement', ascending=False)
    final = pd.concat([data_only, final[final['ASM Name'] == 'TOTAL']], ignore_index=True)


    final['ASM Name'] = final['ASM Name'].astype(str).str.title()

   
    output_path = os.path.join(os.path.dirname(export_file), "Cash_Report.xlsx")
    final.to_excel(output_path, index=False, startrow=1)

 
    wb = load_workbook(output_path)
    ws = wb.active

    chocolate_fill = PatternFill(start_color="7B3F00", end_color="7B3F00", fill_type="solid")
    white_font = Font(color="FFFFFF", bold=True)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    # Header style
    for cell in ws[2]:
        cell.fill = chocolate_fill
        cell.font = white_font

    
    total_row_idx = ws.max_row
    for cell in ws[total_row_idx]:
        cell.fill = chocolate_fill
        cell.font = white_font

   
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.border = thin_border

    
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=4, max_col=4):
        for cell in row:
            cell.number_format = '0%'

    
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
    title_cell = ws.cell(row=1, column=1)
    title_cell.value = "CASH DEPOSIT TARGET"
    title_cell.fill = PatternFill(start_color="00008B", end_color="00008B", fill_type="solid")
    title_cell.font = Font(color="FFFFFF", bold=True)
    title_cell.alignment = Alignment(horizontal='center', vertical='center')

   
    for col_idx, col_cells in enumerate(ws.iter_cols(min_row=3, max_row=ws.max_row), start=1):
        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col_cells)
        col_letter = get_column_letter(col_idx)

        if col_letter == 'A':
            ws.column_dimensions[col_letter].width = min(max_length + 2, 25)
        elif col_letter == 'D':
            ws.column_dimensions[col_letter].width = min(max_length + 2, 14)
        else:
            ws.column_dimensions[col_letter].width = max_length + 2

    wb.save(output_path)
    print("✅ Final Report saved at:", output_path)

except Exception as e:
    print("❌ Error occurred:", e)
