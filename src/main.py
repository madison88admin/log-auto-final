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
    main_data_fields = ['caseNos', 'color', 'unitsCrt', 'totalUnit', 'totalQty', 'carton', 's4Material', 'materialNo']
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
    po_no = first_row.get('sa4PoNo', '')
    ws['D14'] = po_no
    ws['E14'] = f"{po_no} / {po_no}"
    ws['D16'] = po_no
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
        ws.cell(row=row_num, column=21, value=safe_float(row.get('totalCbm', 0)))  # U: TOTAL CBM

    # Place summary after last data row
    summary_row = main_table_start + num_data_rows + 1
    ws.cell(row=summary_row, column=4, value='Summary')
    ws.cell(row=summary_row+1, column=4, value='Total Carton')
    ws.cell(row=summary_row+1, column=5, value=num_data_rows)
    ws.cell(row=summary_row+2, column=4, value='Total Net Weight')
    ws.cell(row=summary_row+2, column=5, value=sum(safe_float(row.get('totalNw', 0)) for row in rows))
    ws.cell(row=summary_row+3, column=4, value='Total Gross Weight')
    ws.cell(row=summary_row+3, column=5, value=sum(safe_float(row.get('totalGw', 0)) for row in rows))
    ws.cell(row=summary_row+4, column=4, value='Total CBM')
    ws.cell(row=summary_row+4, column=5, value=sum(safe_float(row.get('totalCbm', 0)) for row in rows))

    # Color breakdown table (starts at column F)
    color_size_counts = {}
    for row in rows:
        color = row.get('color', '')
        if not color:
            continue
        if color not in color_size_counts:
            color_size_counts[color] = {size: 0 for size in size_names}
        for size in size_names:
            color_size_counts[color][size] += safe_float(row.get(size, 0))
    color_breakdown_row = summary_row + 7
    color_breakdown_col = 6  # Column F
    # Write headers
    ws.cell(row=color_breakdown_row, column=color_breakdown_col, value='Colour')
    for j, size in enumerate(size_names):
        ws.cell(row=color_breakdown_row, column=color_breakdown_col + 1 + j, value=size)
    ws.cell(row=color_breakdown_row, column=color_breakdown_col + 1 + len(size_names), value='Total')
    # Write color breakdown data
    for i, (color, size_dict) in enumerate(color_size_counts.items()):
        ws.cell(row=color_breakdown_row + 1 + i, column=color_breakdown_col, value=color)
        total = 0
        for j, size in enumerate(size_names):
            val = size_dict[size]
            ws.cell(row=color_breakdown_row + 1 + i, column=color_breakdown_col + 1 + j, value=val)
            total += val
        ws.cell(row=color_breakdown_row + 1 + i, column=color_breakdown_col + 1 + len(size_names), value=total)
    # Write total row
    ws.cell(row=color_breakdown_row + 1 + len(color_size_counts), column=color_breakdown_col, value='Total')
    for j, size in enumerate(size_names):
        ws.cell(row=color_breakdown_row + 1 + len(color_size_counts), column=color_breakdown_col + 1 + j, value=sum(size_dict[size] for size_dict in color_size_counts.values()))
    ws.cell(row=color_breakdown_row + 1 + len(color_size_counts), column=color_breakdown_col + 1 + len(size_names), value=sum(sum(size_dict[size] for size in size_names) for size_dict in color_size_counts.values()))

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