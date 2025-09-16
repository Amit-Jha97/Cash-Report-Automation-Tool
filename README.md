import pandas as pd
from tkinter import filedialog, Tk
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# Hide tkinter window
Tk().withdraw()

# Select Export File
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
   
    export_df = pd.read_excel(export_file, sheet_name="Table")
    mapping_df = pd.read_excel(mapping_file, sheet_name="MAPPING")

    # Data preprocessing
    export_df['deposit_date'] = pd.to_datetime(export_df['deposit_date'], errors='coerce')
    export_df = export_df.dropna(subset=['deposit_date'])
    export_df['only_date'] = export_df['deposit_date'].dt.date

    export_df['type'] = export_df['type'].astype(str).str.strip().str.lower()
    export_df['status'] = export_df['status'].astype(str).str.strip().str.lower()

    # Find the latest date
    latest_date = export_df['only_date'].max()
    print("🗓️ Latest Date:", latest_date)

    # Filter export data
    filtered_export = export_df[
        (export_df['only_date'] == latest_date) &
        (export_df['type'].isin(['card', 'cash'])) &
        (export_df['status'] == 'accepted')
    ]

    
    merged = pd.merge(filtered_export, mapping_df, on='code', how='inner')

    mapping_df = mapping_df[mapping_df['target'].fillna(0) != 0]

    
    grouped_amount_asm = (
        merged.groupby('asmname', as_index=False)['amount'].sum()
        .rename(columns={'asmname': 'ASM Name', 'amount': 'Amount'})
    )

   
    unique_target_asm = (
        mapping_df.drop_duplicates(subset='asmname')[['asmname', 'target']]
        .rename(columns={'asmname': 'ASM Name', 'target': 'Target'})
    )

    final_asm = pd.merge(unique_target_asm, grouped_amount_asm, on='ASM Name', how='left')
    final_asm['Amount'] = final_asm['Amount'].fillna(0)
    final_asm['% Achievement'] = final_asm.apply(
        lambda row: (row['Amount'] / row['Target']) if row['Target'] != 0 else 0, axis=1
    )

    # Calculate total for ASM sheet
    total_target_asm = final_asm['Target'].sum()
    total_amount_asm = final_asm['Amount'].sum()
    total_achieved_asm = (total_amount_asm / total_target_asm) if total_target_asm != 0 else 0
    total_row_asm = {
        'ASM Name': 'TOTAL',
        'Target': total_target_asm,
        'Amount': total_amount_asm,
        '% Achievement': total_achieved_asm
    }
    final_asm.loc[len(final_asm)] = total_row_asm
    data_only_asm = final_asm[final_asm['ASM Name'] != 'TOTAL'].sort_values(by='% Achievement', ascending=False)
    final_asm = pd.concat([data_only_asm, final_asm[final_asm['ASM Name'] == 'TOTAL']], ignore_index=True)
    final_asm['ASM Name'] = final_asm['ASM Name'].astype(str).str.title()

    # Calculate and prepare data for HEAD sheet
    grouped_amount_head = (
        merged.groupby('head', as_index=False)['amount'].sum()
        .rename(columns={'head': 'Head', 'amount': 'Amount'})
    )

    # REVISED LOGIC: Pick the unique target value for each ASM and then sum by Head
    unique_asm_targets = mapping_df.drop_duplicates(subset='asmname')[['asmname', 'target']]
    head_asm_map = mapping_df.drop_duplicates(subset='asmname')[['head', 'asmname']]
    merged_targets = pd.merge(head_asm_map, unique_asm_targets, on='asmname', how='inner')
    unique_target_head = (
        merged_targets.groupby('head', as_index=False)['target'].sum()
        .rename(columns={'head': 'Head', 'target': 'Target'})
    )
    
    final_head = pd.merge(unique_target_head, grouped_amount_head, on='Head', how='left')
    final_head['Amount'] = final_head['Amount'].fillna(0)
    final_head['% Achievement'] = final_head.apply(
        lambda row: (row['Amount'] / row['Target']) if row['Target'] != 0 else 0, axis=1
    )

    # Calculate total for HEAD sheet
    total_target_head = final_head['Target'].sum()
    total_amount_head = final_head['Amount'].sum()
    total_achieved_head = (total_amount_head / total_target_head) if total_target_head != 0 else 0
    total_row_head = {
        'Head': 'TOTAL',
        'Target': total_target_head,
        'Amount': total_amount_head,
        '% Achievement': total_achieved_head
    }
    final_head.loc[len(final_head)] = total_row_head
    data_only_head = final_head[final_head['Head'] != 'TOTAL'].sort_values(by='% Achievement', ascending=False)
    final_head = pd.concat([data_only_head, final_head[final_head['Head'] == 'TOTAL']], ignore_index=True)
    final_head['Head'] = final_head['Head'].astype(str).str.title()

    # Create a Pandas Excel writer object to write to multiple sheets
    output_path = os.path.join(os.path.dirname(export_file), "Cash_Report.xlsx")
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        final_asm.to_excel(writer, sheet_name='ASM Report', index=False, startrow=1)
        final_head.to_excel(writer, sheet_name='Head Report', index=False, startrow=1)

    
    wb = load_workbook(output_path)

    
    def style_sheet(ws, title_text, name_col):
        # Define styles
        chocolate_fill = PatternFill(start_color="7B3F00", end_color="7B3F00", fill_type="solid")
        dark_blue_fill = PatternFill(start_color="00008B", end_color="00008B", fill_type="solid")
        white_font = Font(color="FFFFFF", bold=True)
        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        
        # Title styling
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ws.max_column)
        title_cell = ws.cell(row=1, column=1)
        title_cell.value = title_text
        title_cell.fill = dark_blue_fill
        title_cell.font = white_font
        title_cell.alignment = Alignment(horizontal='center', vertical='center')

        # Header style
        for cell in ws[2]:
            cell.fill = chocolate_fill
            cell.font = white_font

        # Total row style
        total_row_idx = ws.max_row
        for cell in ws[total_row_idx]:
            cell.fill = chocolate_fill
            cell.font = white_font

       
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                cell.border = thin_border

        # Percentage format
        for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=4, max_col=4):
            for cell in row:
                cell.number_format = '0%'
        
        # Apply Indian comma format to Target and Amount columns
        indian_format = '[>=10000000]##\,##\,##\,##0;[>=100000]##\,##\,##0;#,##0'
        for col_idx in [2, 3]: 
            for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=col_idx, max_col=col_idx):
                for cell in row:
                    cell.number_format = indian_format

        # Column width adjustment
        for col_idx, col_cells in enumerate(ws.iter_cols(min_row=3, max_row=ws.max_row), start=1):
            max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col_cells)
            col_letter = get_column_letter(col_idx)
            
            if col_letter == 'A':
                ws.column_dimensions[col_letter].width = min(max_length + 2, 25)
            elif col_letter == 'D':
                ws.column_dimensions[col_letter].width = min(max_length + 2, 14)
            else:
                ws.column_dimensions[col_letter].width = max_length + 2

    # Style both sheets
    style_sheet(wb['ASM Report'], "CASH DEPOSIT TARGET (BY ASM)", 'ASM Name')
    style_sheet(wb['Head Report'], "CASH DEPOSIT TARGET (BY HEAD)", 'Head')
    
    # Save the final workbook
    wb.save(output_path)
    print("✅ Final Report saved at:", output_path)

except Exception as e:
    print("❌ Error occurred:", e)
