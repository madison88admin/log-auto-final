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
from openpyxl.styles import Alignment, Font, PatternFill

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
    # Model Name (E16) remains as before
    ws['E16'] = first_row.get('modelName', group_name)
    ws['E7'] = ''

    # Main table starts at row 20 (C20)
    main_table_start = 20
    num_data_rows = len(rows)
    template_data_rows = 1  # Assume template has 1 data row by default
    if num_data_rows > template_data_rows:
        ws.insert_rows(main_table_start + 1, num_data_rows - template_data_rows)

    # Fill columns with correct alignment
    size_names = ['OS', 'XS', 'S', 'M', 'L', 'XL', 'XXL']
    for i, row in enumerate(rows):
        row_num = main_table_start + i
        ws.cell(row=row_num, column=3, value=row.get('caseNos', ''))  # C: CASE NOS
        ws.cell(row=row_num, column=4, value=row.get('color', ''))    # D: COLOR
        ws.cell(row=row_num, column=5, value=row.get('s4Material', ''))  # E: S4 Material
        ws.cell(row=row_num, column=6, value=row.get('materialNo', ''))  # F: Material No#
        ws.cell(row=row_num, column=7, value=row.get('unitsCrt', ''))    # G: Units/CRT (was in N)
        # Sizes H-M
        for j, size in enumerate(size_names):
            ws.cell(row=row_num, column=8+j, value=row.get(size, ''))
        ws.cell(row=row_num, column=14, value=row.get('carton', ''))      # N: Carton (was in K)
        ws.cell(row=row_num, column=15, value=row.get('totalUnit', ''))   # O: Total Unit (was in M)
        ws.cell(row=row_num, column=16, value=safe_float(row.get('totalNw', 0)))  # P: TOTAL N.W.
        ws.cell(row=row_num, column=17, value=safe_float(row.get('totalGw', 0)))  # Q: TOTAL G.W.
        ws.cell(row=row_num, column=18, value=row.get('Length', ''))    # R: Length
        ws.cell(row=row_num, column=19, value=row.get('Width', ''))     # S: Width
        ws.cell(row=row_num, column=20, value=row.get('Height', ''))    # T: Height
        cbm_value = safe_float(row.get('cbm', row.get('totalCbm', 0)))
        ws.cell(row=row_num, column=21, value=cbm_value)  # U: CBM
        ws.cell(row=row_num, column=22, value=cbm_value)  # V: TOTAL CBM

    # Place summary and color breakdown headers on the same row after the main table, with a gap
    gap = 1
    header_row = main_table_start + num_data_rows + gap + 1
    summary_col = 4  # D
    value_col = 5    # E
    color_breakdown_col = 6  # F

    # Merge and style the summary header (D/E)
    ws.merge_cells(start_row=header_row, start_column=summary_col, end_row=header_row, end_column=value_col)
    header_cell = ws.cell(row=header_row, column=summary_col)
    header_cell.value = 'Summary'
    header_cell.alignment = Alignment(horizontal='center', vertical='center')
    header_cell.font = Font(bold=True)
    header_cell.fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

    # (Assume color breakdown headers are present in the template at header_row, columns F onward)

    # Write summary values (labels and values) below the summary header
    ws.cell(row=header_row+1, column=summary_col, value='Total Carton')
    ws.cell(row=header_row+1, column=value_col, value=num_data_rows)
    ws.cell(row=header_row+2, column=summary_col, value='Total Net Weight (kg)')
    ws.cell(row=header_row+2, column=value_col, value=round(sum(safe_float(row.get('totalNw', 0)) for row in rows), 3))
    ws.cell(row=header_row+3, column=summary_col, value='Total Gross Weight (kg)')
    ws.cell(row=header_row+3, column=value_col, value=round(sum(safe_float(row.get('totalGw', 0)) for row in rows), 3))
    ws.cell(row=header_row+4, column=summary_col, value='Total CBM')
    ws.cell(row=header_row+4, column=value_col, value=round(sum(safe_float(row.get('cbm', row.get('totalCbm', 0))) for row in rows), 3))

    # Write color breakdown data below the color breakdown headers
    color_breakdown_data_row = header_row + 1
    # Calculate color_size_counts before using it
    color_size_counts = {}
    for row in rows:
        color = row.get('color', '')
        if not color:
            continue
        if color not in color_size_counts:
            color_size_counts[color] = {size: 0 for size in size_names}
        for size in size_names:
            color_size_counts[color][size] += safe_float(row.get(size, 0))

    for i, (color, size_dict) in enumerate(color_size_counts.items()):
        ws.cell(row=color_breakdown_data_row + i, column=color_breakdown_col, value=color)
        total = 0
        for j, size in enumerate(size_names):
            val = size_dict[size]
            ws.cell(row=color_breakdown_data_row + i, column=color_breakdown_col + 1 + j, value=val)
            total += val
        ws.cell(row=color_breakdown_data_row + i, column=color_breakdown_col + 1 + len(size_names), value=total)

    # Write total row for color breakdown
    ws.cell(row=color_breakdown_data_row + len(color_size_counts), column=color_breakdown_col, value='Total')
    for j, size in enumerate(size_names):
        ws.cell(row=color_breakdown_data_row + len(color_size_counts), column=color_breakdown_col + 1 + j, value=sum(size_dict[size] for size_dict in color_size_counts.values()))
    ws.cell(row=color_breakdown_data_row + len(color_size_counts), column=color_breakdown_col + 1 + len(size_names), value=sum(sum(size_dict[size] for size in size_names) for size_dict in color_size_counts.values()))

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