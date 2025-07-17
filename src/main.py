import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from copy import copy
from fastapi import FastAPI, File, UploadFile, Form
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
import io
import re
from typing import List
import json
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.cell.cell import MergedCell

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Canonical header mapping (adjust as needed to match frontend)
HEADER_MAPPING = {
    'CASE NOS': 'caseNos',
    'SA4 PO NO#': 'sa4PoNo',
    'CARTON PO NO#': 'cartonPoNo',
    'SAP STYLE NO': 'sapStyleNo',
    'STYLE NAME #': 'modelName',
    'S4 Material': 's4Material',
    'Material No#': 'materialNo',
    'COLOR': 'color',
    'Size': 'size',
    'Total QTY': 'totalQty',
    'CARTON': 'carton',
    'QTY / CARTON': 'unitsCrt',
    'TOTAL\nQTY': 'totalUnit',
    'N.W / ctn.': 'nwCtn',
    'TOTAL\nN.W.': 'totalNw',
    'G.W. / ctn': 'gwCtn',
    'TOTAL\nG.W.': 'totalGw',
    'MEAS.\nCM': 'measCm',
    'TOTAL\nCBM': 'totalCbm',
    'OS': 'OS', 'XS': 'XS', 'S': 'S', 'M': 'M', 'L': 'L', 'XL': 'XL', 'XXL': 'XXL',
    # Add any other variants as needed
}

# Choose the grouping key (adjust as needed)
GROUP_KEY = 'sa4PoNo'


# Normalize header
def normalize_header(header):
    return header.strip().replace('\n', ' ').replace('\r', '').replace('  ', ' ')

def extract_tables(file_bytes):
    xls = pd.ExcelFile(file_bytes)
    tables = []
    # Updated main_data_fields to match the main table headers
    main_data_fields = [
        'cartonNo', 'color', 's4Material', 'materialNo',
        'OS', 'XS', 'S', 'M', 'L', 'XL', 'XXL',
        'unitsCrt', 'totalUnit', 'totalNw', 'totalGw',
        'carton', 'Length', 'Width', 'Height',
        'cbm', 'totalCbm'
    ]
    for sheet_name in xls.sheet_names:
        df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None, dtype=str)
        header_indices = [i for i, row in df_raw.iterrows() if str(row.iloc[0]).strip() == "CASE NOS" or str(row.iloc[0]).strip() == "Carton#"]
        for idx, header_idx in enumerate(header_indices):
            start = header_idx
            end = header_indices[idx + 1] if idx + 1 < len(header_indices) else len(df_raw)
            table_rows = df_raw.iloc[start:end].values.tolist()
            if not table_rows or len(table_rows) < 2:
                continue
            header_row = [str(cell).strip() if cell is not None else "" for cell in table_rows[0]]
            mapped_keys = [HEADER_MAPPING.get(str(h).strip(), str(h).strip()) for h in header_row]
            print('Header row:', header_row)
            print('Mapped keys:', mapped_keys)
            table_data = []
            for data_row in table_rows[1:]:
                # Skip rows that are completely empty
                if all((cell is None or str(cell).strip() == "" or str(cell).lower() == "nan") for cell in data_row):
                    continue
                row_dict = {mapped_keys[i]: data_row[i] if i < len(data_row) else "" for i in range(len(header_row))}
                # Assign Length, Width, Height from R, S, T columns (Excel columns 18, 19, 20; Python indices 17, 18, 19)
                if len(data_row) > 17:
                    row_dict['Length'] = data_row[17]
                if len(data_row) > 18:
                    row_dict['Width'] = data_row[18]
                if len(data_row) > 19:
                    row_dict['Height'] = data_row[19]
                # Only include rows with at least one main data field filled
                if any(str(row_dict.get(f, '')).strip() not in ['', 'nan', 'None'] for f in main_data_fields):
                    table_data.append(row_dict)
            if table_data:
                print('Sample data row:', table_data[0])
            if not table_data:
                continue
            sheet_name = next((row.get('sa4PoNo', '') or row.get('cartonNo', '') for row in table_data if row.get('sa4PoNo', '') or row.get('cartonNo', '')), f'Table{len(tables)+1}')
            safe_name = re.sub(r'[:\\/?*\[\]]', '_', str(sheet_name))[:31]
            tables.append({'rows': table_data, 'sheet_name': safe_name})
    return tables

def copy_worksheet(template_ws, target_ws):
    for row in template_ws.iter_rows():
        for cell in row:
            if cell.__class__.__name__ == "MergedCell":
                continue
            new_cell = target_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.border = copy(cell.border)
                new_cell.fill = copy(cell.fill)
                new_cell.number_format = copy(cell.number_format)
                new_cell.protection = copy(cell.protection)
                new_cell.alignment = copy(cell.alignment)
    for col, dim in template_ws.column_dimensions.items():
        target_ws.column_dimensions[col].width = dim.width
    for row, dim in template_ws.row_dimensions.items():
        target_ws.row_dimensions[row].height = dim.height
    for merged_range in template_ws.merged_cells.ranges:
        target_ws.merge_cells(str(merged_range))

def safe_float(val):
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0

def fill_template_with_data(ws, rows, group_name):
    if not rows:
        return
    # Fill header fields (adjust cell addresses as needed)
    first_row = rows[0]
    s4_hana_sku = str(first_row.get('s4Material', '') or first_row.get('E', ''))
    po_line_value = s4_hana_sku[:-2] if len(s4_hana_sku) > 2 else s4_hana_sku
    # Set PO-Line (D14) and Model # (D16) to trimmed S4 HANA SKU
    ws['D14'] = po_line_value
    ws['D16'] = po_line_value
    # Set SAP PO# (E14) and PO NO# to sheet name
    ws['E14'] = group_name
    # --- Model Name logic: match frontend (App.tsx) ---
    model_name = ''
    for i in range(len(rows)):
        case_nos_val = str(rows[i].get('caseNos', '')).lower()
        if 'case' in case_nos_val:
            if i - 2 >= 0:
                model_name = rows[i-2].get('caseNos', '')
            break
    # Prefer the modelName field if available
    if rows and rows[0].get('modelName'):
        model_name = rows[0]['modelName']

    # Set Model Name (E16) in the worksheet
    ws['E16'] = model_name
    ws['E7'] = ''

    # Main table starts at row 20 (C20)
    main_table_start = 20
    num_data_rows = len(rows)
    template_data_rows = 1  # Assume template has 1 data row by default

    # --- Unmerge any merged cells in the main table data area (C20:V<end>) BEFORE inserting rows ---
    main_table_data_start = main_table_start
    main_table_data_end = main_table_start + num_data_rows + 20  # a safe buffer
    for merged_range in list(ws.merged_cells.ranges):
        if (
            merged_range.min_row >= main_table_data_start and
            merged_range.max_row <= main_table_data_end and
            merged_range.min_col >= 3 and
            merged_range.max_col <= 22
        ):
            ws.unmerge_cells(str(merged_range))

    if num_data_rows > template_data_rows:
        ws.insert_rows(main_table_start + 1, num_data_rows - template_data_rows)

    # --- MAIN TABLE: Propagate carton numbers and track ranges for merging (match App.tsx logic) ---
    effective_carton_nos = []
    last_carton_no = ''
    carton_row_ranges = {}
    for i, row in enumerate(rows):
        carton_no = str(row.get('caseNos', '')).strip()
        # Only propagate and track valid carton numbers (not empty, not 'nan')
        if carton_no and carton_no.lower() != 'nan':
            last_carton_no = carton_no
        effective_carton_nos.append(last_carton_no)
        row_num = main_table_start + i
        if last_carton_no and last_carton_no.lower() != 'nan':
            if last_carton_no not in carton_row_ranges:
                carton_row_ranges[last_carton_no] = {'start': row_num, 'end': row_num}
            else:
                carton_row_ranges[last_carton_no]['end'] = row_num

    # --- Write main table with split-carton OS logic and inline copy-down for D/E/F ---
    prev_color = ''
    prev_s4 = ''
    prev_ecc = ''
    for i, row in enumerate(rows):
        row_num = main_table_start + i
        carton_no = effective_carton_nos[i]
        # Only write Carton# for the first row in the group, else leave blank
        if carton_no and carton_no.lower() != 'nan' and carton_row_ranges.get(carton_no, {}).get('start') == row_num:
            ws.cell(row=row_num, column=3, value=carton_no)  # C: CASE NOS
        else:
            ws.cell(row=row_num, column=3, value='')

        # Color (D)
        color = str(row.get('color', '')).strip()
        if not color or color.lower() in ['nan', 'none']:
            color = prev_color
        ws.cell(row=row_num, column=4, value=color)
        prev_color = color

        # S4 HANA SKU (E)
        s4sku = str(row.get('s4Material', '')).strip()
        if not s4sku or s4sku.lower() in ['nan', 'none']:
            s4sku = prev_s4
        ws.cell(row=row_num, column=5, value=s4sku)
        prev_s4 = s4sku

        # ECC Material No (F)
        ecc = str(row.get('materialNo', '')).strip()
        if not ecc or ecc.lower() in ['nan', 'none']:
            ecc = prev_ecc
        ws.cell(row=row_num, column=6, value=ecc)
        prev_ecc = ecc

        # --- SPLIT CARTON LOGIC FOR OS COLUMN (split carton -> J, else L) ---
        if carton_no and carton_no.lower() != 'nan' and sum(1 for c in effective_carton_nos if c == carton_no) > 1:
            os_val = row.get('totalQty', '')  # Column J (index 9)
        else:
            os_val = row.get('unitsCrt', '')  # Column L (index 11)
        ws.cell(row=row_num, column=7, value=os_val)    # G: OS

        # Sizes H-M
        size_names = ['OS', 'XS', 'S', 'M', 'L', 'XL', 'XXL']
        for j, size in enumerate(size_names[1:]):  # skip OS, already filled
            ws.cell(row=row_num, column=8+j, value=row.get(size, ''))

        ws.cell(row=row_num, column=14, value=row.get('carton', ''))      # N: Carton
        ws.cell(row=row_num, column=15, value=row.get('totalUnit', ''))   # O: Total Unit
        ws.cell(row=row_num, column=16, value=safe_float(row.get('totalNw', 0)))  # P: TOTAL N.W.
        ws.cell(row=row_num, column=17, value=safe_float(row.get('totalGw', 0)))  # Q: TOTAL G.W.
        ws.cell(row=row_num, column=18, value=row.get('Length', ''))    # R: Length
        ws.cell(row=row_num, column=19, value=row.get('Width', ''))     # S: Width
        ws.cell(row=row_num, column=20, value=row.get('Height', ''))    # T: Height
        cbm_value = safe_float(row.get('cbm', row.get('totalCbm', 0)))
        ws.cell(row=row_num, column=21, value=cbm_value)  # U: CBM
        ws.cell(row=row_num, column=22, value=cbm_value)  # V: TOTAL CBM

    # --- Only merge Carton# and N–V columns (do not merge D/E/F) ---
    for carton_no, rng in carton_row_ranges.items():
        if carton_no and carton_no.lower() != 'nan' and rng['end'] > rng['start']:
            ws.merge_cells(start_row=rng['start'], start_column=3, end_row=rng['end'], end_column=3)
            for col in range(14, 23):  # N–V
                ws.merge_cells(start_row=rng['start'], start_column=col, end_row=rng['end'], end_column=col)

    # --- SUMMARY AND COLOR BREAKDOWN (replicate frontend logic) ---
    # Always write summary and color breakdown at fixed positions after the main table
    size_names = ['OS', 'XS', 'S', 'M', 'L', 'XL', 'XXL']
    main_table_start = 20
    num_data_rows = len(rows)
    summary_start_row = main_table_start + num_data_rows + 1
    summary_col = 4  # D
    value_col = 5    # E
    color_breakdown_col = 6  # F

    # --- CLEAR TEMPLATE SUMMARY AND COLOR BREAKDOWN AREA ---

    # Determine where the main table ends
    main_table_end_row = main_table_start + num_data_rows - 1
    clear_start_row = main_table_end_row + 1
    clear_end_row = summary_start_row + 20
    for row in range(clear_start_row, clear_end_row):
        for col in range(4, 15):  # D to N (or further if needed)
            cell = ws.cell(row=row, column=col)
            if isinstance(cell, MergedCell):
                continue
            cell.value = None
            cell.font = None
            cell.alignment = None
            cell.fill = PatternFill()
            cell.border = Border()

    # Unmerge any merged cells in this area
    for merged_range in list(ws.merged_cells.ranges):
        if (merged_range.min_row >= clear_start_row and
            merged_range.max_row <= clear_end_row and
            merged_range.min_col >= 4 and
            merged_range.max_col <= 15):
            ws.unmerge_cells(str(merged_range))

    # --- SUMMARY SECTION ---image.png
    # Merge D and E for the summary name
    ws.merge_cells(start_row=summary_start_row, start_column=summary_col, end_row=summary_start_row, end_column=value_col)
    header_cell = ws.cell(row=summary_start_row, column=summary_col)
    header_cell.value = "Summary"
    header_cell.alignment = Alignment(horizontal='center', vertical='center')
    header_cell.font = Font(bold=True)
    header_cell.fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

    # Write summary values (labels and values) below the summary header
    ws.cell(row=summary_start_row+1, column=summary_col, value='Total Carton')
    ws.cell(row=summary_start_row+1, column=value_col, value=num_data_rows)
    ws.cell(row=summary_start_row+2, column=summary_col, value='Total Net Weight (kg)')
    ws.cell(row=summary_start_row+2, column=value_col, value=round(sum(safe_float(row.get('totalNw', 0)) for row in rows), 3))
    ws.cell(row=summary_start_row+3, column=summary_col, value='Total Gross Weight (kg)')
    ws.cell(row=summary_start_row+3, column=value_col, value=round(sum(safe_float(row.get('totalGw', 0)) for row in rows), 3))
    ws.cell(row=summary_start_row+4, column=summary_col, value='Total CBM')
    ws.cell(row=summary_start_row+4, column=value_col, value=round(sum(safe_float(row.get('cbm', row.get('totalCbm', 0))) for row in rows),3))

    # --- COLOR BREAKDOWN SECTION (replicating frontend split-carton logic) ---
    # Place color breakdown headers to align with the summary header row
    color_breakdown_start_row = summary_start_row  # Align with summary header row
    color_breakdown_start_col = color_breakdown_col
    color_headers = ['Color'] + size_names + ['Total']
    border_style = Side(style='thin')
    header_border = Border(top=border_style, left=border_style, bottom=border_style, right=border_style)

    # Write headers with borders
    for i, header in enumerate(color_headers):
        cell = ws.cell(row=color_breakdown_start_row, column=color_breakdown_start_col + i, value=header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = header_border

    # --- NEW LOGIC: propagate carton numbers and build carton count map ---
    effective_carton_nos = []
    last_carton_no = ''
    for row in rows:
        carton_no = str(row.get('caseNos', '')).strip()
        if carton_no:
            last_carton_no = carton_no
        effective_carton_nos.append(last_carton_no)
    carton_count_map = {}
    for carton_no in effective_carton_nos:
        if not carton_no:
            continue
        carton_count_map[carton_no] = carton_count_map.get(carton_no, 0) + 1

    # Calculate color breakdown data with split-carton logic for OS
    color_map = {}
    for i, row in enumerate(rows):
        color = str(row.get('color', '')).strip()
        if not color or color.lower() in ['nan', 'none']:
            continue
        if color not in color_map:
            color_map[color] = [0] * len(size_names)
        carton_no = effective_carton_nos[i]
        # OS column (index 0): use split carton logic
        if carton_no and carton_count_map[carton_no] > 1:
            # Split carton: use OS value (unitsCrt)
            os_val = safe_float(row.get('OS', row.get('unitsCrt', 0)))
            color_map[color][0] += os_val
        else:
            # Single carton: use Total Unit value
            os_val = safe_float(row.get('totalUnit', 0))
            color_map[color][0] += os_val
        # Other sizes: sum as before
        for j, size in enumerate(size_names[1:], 1):
            size_val = safe_float(row.get(size, 0))
            color_map[color][j] += size_val

    valid_colors = [color for color in color_map.keys() if color and color.strip().lower() not in ['nan', 'none']]

    # Write color breakdown rows dynamically with borders
    for color_idx, color in enumerate(valid_colors):
        row_num = color_breakdown_start_row + 1 + color_idx
        color_cell = ws.cell(row=row_num, column=color_breakdown_start_col, value=color)
        color_cell.alignment = Alignment(horizontal='left')
        color_cell.border = header_border
        for i in range(len(size_names)):
            cell = ws.cell(row=row_num, column=color_breakdown_start_col + 1 + i, value=color_map[color][i])
            cell.alignment = Alignment(horizontal='center')
            cell.border = header_border
        total_value = sum(color_map[color])
        total_cell = ws.cell(row=row_num, column=color_breakdown_start_col + 8, value=total_value)
        total_cell.font = Font(bold=True)
        total_cell.alignment = Alignment(horizontal='center')
        total_cell.border = header_border

    # Write the total row for color breakdown with borders
    total_row_num = color_breakdown_start_row + 1 + len(valid_colors)
    total_label_cell = ws.cell(row=total_row_num, column=color_breakdown_start_col, value="Total")
    total_label_cell.font = Font(bold=True)
    total_label_cell.alignment = Alignment(horizontal='left')
    total_label_cell.border = header_border
    for i in range(len(size_names)):
        size_total = sum(color_map[color][i] for color in valid_colors)
        cell = ws.cell(row=total_row_num, column=color_breakdown_start_col + 1 + i, value=size_total)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
        cell.border = header_border
    grand_total = sum(sum(color_map[color]) for color in valid_colors)
    grand_total_cell = ws.cell(row=total_row_num, column=color_breakdown_start_col + 8, value=grand_total)
    grand_total_cell.font = Font(bold=True)
    grand_total_cell.alignment = Alignment(horizontal='center')
    grand_total_cell.border = header_border

    # --- Dynamically extend the main report border to the last row of the color breakdown table ---
    leftmost_col = color_breakdown_start_col
    rightmost_col = color_breakdown_start_col + len(color_headers) - 1
    thick_border = Border(bottom=Side(style='thick'))
    for col in range(leftmost_col, rightmost_col + 1):
        cell = ws.cell(row=total_row_num, column=col)
        # Start with the thick bottom border
        border = Border(
            left=cell.border.left,
            right=cell.border.right,
            top=cell.border.top,
            bottom=thick_border.bottom
        )
        # Add thick left border to the first column
        if col == leftmost_col:
            border = Border(
                left=Side(style='thick'),
                right=border.right,
                top=border.top,
                bottom=border.bottom
            )
        # Add thick right border to the last column
        if col == rightmost_col:
            border = Border(
                left=border.left,
                right=Side(style='thick'),
                top=border.top,
                bottom=border.bottom
            )
        cell.border = border

    # --- Clear all borders in the top section (B2:W16) ---
    for row in range(2, 17):  # 2 to 16 inclusive
        for col in range(2, 24):  # B (2) to W (23)
            ws.cell(row=row, column=col).border = Border()

    # --- Apply only the requested border ---
    # Double border for cell C3 (Packing List)
    ws['C3'].border = Border(
        left=Side(style='double'),
        right=Side(style='double'),
        top=Side(style='double'),
        bottom=Side(style='double')
    )

    # --- Apply only the requested borders ---
    # Double border for cell C3 (Packing List)
    ws['C3'].border = Border(
        left=Side(style='double'),
        right=Side(style='double'),
        top=Side(style='double'),
        bottom=Side(style='double')
    )

    # --- Apply thin borders to the main table (C20:V<end>) ---
    thin_side = Side(style='thin')
    thin_border = Border(top=thin_side, left=thin_side, right=thin_side, bottom=thin_side)
    main_table_end_row = main_table_start + num_data_rows - 1
    for row in range(main_table_start, main_table_end_row + 1):
        for col in range(3, 23):  # C (3) to V (22)
            ws.cell(row=row, column=col).border = thin_border

    # --- Apply thin borders to the summary table (D/E, summary_start_row+1 to summary_start_row+4) ---
    for row in range(summary_start_row + 1, summary_start_row + 5):
        for col in range(4, 6):  # D (4) and E (5)
            ws.cell(row=row, column=col).border = thin_border

    # --- Apply thin borders to the color breakdown table (F to O, color_breakdown_start_row to total_row_num) ---
    for row in range(color_breakdown_start_row, total_row_num + 1):
        for col in range(6, 15):  # F (6) to O (14)
            ws.cell(row=row, column=col).border = thin_border

    # --- Draw a robust thick outer border around the entire report area ---
    top_row = 2
    bottom_row = total_row_num + 1  # last row of color breakdown
    left_col = 2  # Column B
    right_col = 23  # Column W

    # Top and bottom borders
    for col in range(left_col, right_col + 1):
        ws.cell(row=top_row, column=col).border = Border(
            top=Side(style='thick'),
            left=ws.cell(row=top_row, column=col).border.left,
            right=ws.cell(row=top_row, column=col).border.right,
            bottom=ws.cell(row=top_row, column=col).border.bottom
        )
        ws.cell(row=bottom_row, column=col).border = Border(
            bottom=Side(style='thick'),
            left=ws.cell(row=bottom_row, column=col).border.left,
            right=ws.cell(row=bottom_row, column=col).border.right,
            top=ws.cell(row=bottom_row, column=col).border.top
        )

    # Left and right borders
    for row in range(top_row, bottom_row + 1):
        ws.cell(row=row, column=left_col).border = Border(
            left=Side(style='thick'),
            top=ws.cell(row=row, column=left_col).border.top,
            right=ws.cell(row=row, column=left_col).border.right,
            bottom=ws.cell(row=row, column=left_col).border.bottom
        )
        ws.cell(row=row, column=right_col).border = Border(
            right=Side(style='thick'),
            top=ws.cell(row=row, column=right_col).border.top,
            left=ws.cell(row=row, column=right_col).border.left,
            bottom=ws.cell(row=row, column=right_col).border.bottom
        )
        
    robust_copy_down(rows, ['color', 's4Material', 'materialNo'])

def robust_copy_down(rows, keys):
    last_values = {k: '' for k in keys}
    for row in rows:
        for k in keys:
            val = str(row.get(k, '')).strip()
            if not val or val.lower() in ['nan', 'none']:
                row[k] = last_values[k]
            else:
                last_values[k] = val

# In generate_reports, skip NaN/Unknown group keys
@app.post("/generate-reports/")
async def generate_reports(
    file: UploadFile = File(...),
    table_keys: str = Form(None)  # JSON stringified list of keys
):
    print("LOG TEST: /generate-reports endpoint called")
    file_bytes = await file.read()
    tables = extract_tables(io.BytesIO(file_bytes))
    template_wb = load_workbook("ReportTemplate.xlsx")
    template_ws = template_wb['Report']
    combined_wb = Workbook()
    if combined_wb.active is not None and combined_wb.active.title == "Sheet":
        combined_wb.remove(combined_wb.active)
    requested_keys = json.loads(table_keys) if table_keys else [t['sheet_name'] for t in tables]
    print("Available tables:", [t['sheet_name'] for t in tables])
    print("Requested keys:", requested_keys)
    normalized_tables = {str(t['sheet_name']).strip().lower(): t for t in tables}
    for key in requested_keys:
        if not key or str(key).lower() in ['nan', 'unknown']:
            continue
        norm_key = str(key).strip().lower()
        table = normalized_tables.get(norm_key)
        if not table or not table['rows']:
            continue
        ws = combined_wb.create_sheet(title=table['sheet_name'])
        copy_worksheet(template_ws, ws)
        fill_template_with_data(ws, table['rows'], table['sheet_name'])
    output = io.BytesIO()
    combined_wb.save(output)
    output.seek(0)
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename=AllReports.xlsx"}
    )