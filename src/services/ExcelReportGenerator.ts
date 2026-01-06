import * as XLSX from 'xlsx';
import { OrderReport, ColorSummary, TKListData } from '../types';

const TEMPLATE_URL = '/ReportTemplate.xlsx';

async function loadTemplateWorkbook(): Promise<XLSX.WorkBook> {
  const response = await fetch(TEMPLATE_URL);
  if (!response.ok) throw new Error('Failed to fetch template');
  const arrayBuffer = await response.arrayBuffer();
  return XLSX.read(arrayBuffer, { type: 'array' });
}

export class ExcelReportGenerator {
  // Minimal sanity test for Excel generation
  static async sanityTestExcel(): Promise<ArrayBuffer> {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([
      ['Hello', 'World'],
      [1, 2],
      [3, 4]
    ]);
    XLSX.utils.book_append_sheet(wb, ws, 'Test');
    return XLSX.write(wb, { bookType: 'xlsx', type: 'array', compression: true });
  }

  private async loadTemplateSheet(): Promise<XLSX.WorkSheet> {
    try {
      const templateWb = await loadTemplateWorkbook();
      const templateSheet = templateWb.Sheets['Report'];
      // Deep copy the sheet to avoid mutating the original
      return JSON.parse(JSON.stringify(templateSheet));
    } catch (e) {
      // Fallback: return a minimal sheet if template fails
      console.error('Failed to load template, using fallback sheet:', e);
      return XLSX.utils.aoa_to_sheet([
        ['Fallback', 'Sheet'],
        ['No template found', '']
      ]);
    }
  }

  private populateColorData(ws: XLSX.WorkSheet, orderReport: OrderReport, startRow: number): number {
    let currentRow = startRow;
    // Debug output
    console.log('Populating color data:', orderReport.colorSummary);
    // Add each color with its quantity
    orderReport.colorSummary.forEach((colorData, index) => {
      // Calculate column G (O divided by N) for color data
      const totalQty = colorData.quantity;
      const weightPerCarton = 1; // Default weight per carton for color summary
      let gValue: number | string = '';
      if (totalQty && weightPerCarton) {
        gValue = totalQty / weightPerCarton;
      } else if (index > 0) {
        // Copy from above if current values are invalid
        const prevCell = ws[XLSX.utils.encode_cell({ r: currentRow - 1, c: 6 })];
        gValue = prevCell ? prevCell.v : '';
      }
      
      const rowData = [
        index + 1,                    // A: Carton# (sequential)
        colorData.color,              // B: Color
        '',                          // C: S4 HANA SKU
        '',                          // D: ECC Material No
        colorData.quantity,          // E: QS (quantity)
        '',                          // F: XS
        gValue,                      // G: S (calculated)
        '',                          // H: M
        '',                          // I: L
        '',                          // J: XL
        '',                          // K: XXL
        '',                          // L: Unit/CHT
        colorData.quantity,          // M: Total Unit/Qty
        weightPerCarton,             // N: Weight/Carton
        colorData.quantity,          // O: Total
        '',                          // P: Carton Size
        '',                          // Q: 
        '',                          // R: CBM
        ''                           // S: Total CBM
      ];
      // Convert row data to worksheet cells
      rowData.forEach((value, colIndex) => {
        const cellAddress = XLSX.utils.encode_cell({ r: currentRow, c: colIndex });
        ws[cellAddress] = { v: value, t: typeof value === 'number' ? 'n' : 's' };
      });
      currentRow++;
    });
    // Add total row
    // Calculate column G for total row
    const totalWeightPerCarton = 1; // Default weight per carton for total
    const totalGValue = orderReport.totalQuantity && totalWeightPerCarton ? 
      orderReport.totalQuantity / totalWeightPerCarton : '';
    
    const totalRowData = [
      '',                           // A: Carton#
      'Total',                      // B: Color
      '',                          // C: S4 HANA SKU
      '',                          // D: ECC Material No
      orderReport.totalQuantity,   // E: QS (total quantity)
      '',                          // F: XS
      totalGValue,                 // G: S (calculated)
      '',                          // H: M
      '',                          // I: L
      '',                          // J: XL
      '',                          // K: XXL
      '',                          // L: Unit/CHT
      orderReport.totalQuantity,   // M: Total Unit/Qty
      totalWeightPerCarton,        // N: Weight/Carton
      orderReport.totalQuantity,   // O: Total
      '',                          // P: Carton Size
      '',                          // Q: 
      '',                          // R: CBM
      ''                           // S: Total CBM
    ];
    totalRowData.forEach((value, colIndex) => {
      const cellAddress = XLSX.utils.encode_cell({ r: currentRow, c: colIndex });
      ws[cellAddress] = { v: value, t: typeof value === 'number' ? 'n' : 's' };
    });
    // Debug output
    console.log('Worksheet after color data:', ws);
    return currentRow + 1;
  }

  async generateExcelReport(orderReports: OrderReport[], tkListData: TKListData, originalFileName: string): Promise<ArrayBuffer> {
    // Create a new workbook
    const workbook = XLSX.utils.book_new();
    // Create Menu sheet
    const menuWs = XLSX.utils.aoa_to_sheet([
      ['PACKING LIST REPORTS'],
      ['Generated from: ' + originalFileName],
      ['Generated on: ' + new Date().toLocaleString()],
      [],
      ['Available Reports:'],
      ...orderReports.map((report, index) => [
        `${index + 1}. Order ${report.orderNumber} (${report.colorSummary.length} colors, ${report.totalQuantity} total qty)`
      ])
    ]);
    XLSX.utils.book_append_sheet(workbook, menuWs, 'Menu');

    // Debug: log all order numbers
    console.log('Generating sheets for orders:', orderReports.map(r => r.orderNumber));
    let sheetCount = 0;
    // Create a sheet for each order report using the template
    for (const orderReport of orderReports) {
      const ws = await this.loadTemplateSheet();
      // Map all fields from TK List data to template
      this.mapTKListDataToTemplate(ws, orderReport, tkListData);
      // Debug output
      console.log('Mapping TK List data to template for order:', orderReport.orderNumber);
      // Populate color data starting from row 20 (index 19 or 20 depending on template)
      this.populateColorData(ws, orderReport, 19);
      // Calculate worksheet range (!ref)
      let maxRow = 19 + (tkListData.orderGroups[orderReport.orderNumber]?.length || 1) + 10; // +10 for summary
      let maxCol = 11; // Assume 11 columns (A-K)
      ws['!ref'] = XLSX.utils.encode_range({ s: { r: 19, c: 0 }, e: { r: maxRow, c: maxCol } });
      // Clean sheet name for Excel compatibility
      const sheetName = orderReport.orderNumber.replace(/[\/\\?*\[\]]/g, '_').substring(0, 31);
      XLSX.utils.book_append_sheet(workbook, ws, sheetName);
      sheetCount++;
      // Debug: log data written to this sheet
      console.log(`Sheet generated: ${sheetName}, rows: ${maxRow - 19 + 1}, columns: ${maxCol + 1}`);
    }
    // Debug: log total sheets
    console.log('Total sheets generated (excluding menu):', sheetCount);
    // Debug output
    console.log('Workbook after all sheets:', workbook);
    // Generate Excel buffer
    const excelBuffer = XLSX.write(workbook, { 
      bookType: 'xlsx', 
      type: 'array',
      compression: true 
    });
    return excelBuffer;
  }

  private mapTKListDataToTemplate(ws: XLSX.WorkSheet, orderReport: OrderReport, tkListData: TKListData): void {
    // Debug output
    console.log('Mapping data for order:', orderReport.orderNumber, tkListData);
    const orderData = tkListData.orderGroups[orderReport.orderNumber] || [];
    if (orderData.length === 0) return;

    // --- PATCH: Write PO-Line, SAP PO#, Model #, Model Name ---
    // Adjust these cell addresses as needed for your template
    // PO-Line (B2)
    ws['B2'] = { v: orderData[0]?.sa4PoNo || '', t: 's' };
    // SAP PO# (C2) - format as SA4 PO NO# / SA4 PO NO#
    const poNo = orderData[0]?.sa4PoNo || '';
    ws['C2'] = { v: poNo ? `${poNo} / ${poNo}` : '', t: 's' };
    // Model # (B3)
    ws['B3'] = { v: orderData[0]?.sa4PoNo || '', t: 's' };
    // Model Name (C3)
    ws['C3'] = { v: orderReport.modelName || '', t: 's' };
    // --- END PATCH ---

    // 1. Map main table fields
    // Find the starting row for the data table in the template (assume header is at row 19, data starts at 20)
    let startRow = 19; // 0-indexed (row 20 in Excel)
    orderData.forEach((row, i) => {
      // Map each field to the correct column in the template
      // Carton# (Case Nos) - Col 0
      ws[XLSX.utils.encode_cell({ r: startRow + i, c: 0 })] = { v: row['caseNos'], t: 's' };
      // SA4 PO NO# (PO-Line/Model/SAP PO#) - Col 1
      ws[XLSX.utils.encode_cell({ r: startRow + i, c: 1 })] = { v: row['sa4PoNo'], t: 's' };
      // S4 HANA SKU (S4 Material) - Col 2
      ws[XLSX.utils.encode_cell({ r: startRow + i, c: 2 })] = { v: row['s4Material'], t: 's' };
      // ECC Material No (Material No#) - Col 3
      ws[XLSX.utils.encode_cell({ r: startRow + i, c: 3 })] = { v: row['materialNo'], t: 's' };
      // Color (COLOR) - Col 4
      ws[XLSX.utils.encode_cell({ r: startRow + i, c: 4 })] = { v: row['color'], t: 's' };
      // Units/CRT (CARTON) - Col 5
      ws[XLSX.utils.encode_cell({ r: startRow + i, c: 5 })] = { v: row['carton'], t: 'n' };
      // Total Unit (TOTAL QTY) - Col 6 (G)
      const totalQty = row['totalQty'] ?? row['totalQty2'];
      const weightPerCarton = row['totalNw']; // Column N (13)

      // Calculate column G (O divided by N) - Col 6 (G)
      let gValue: number | string = '';
      if (totalQty && weightPerCarton && weightPerCarton !== 0) {
        gValue = totalQty / weightPerCarton;
      } else if (i > 0) {
        // Copy from above if current values are invalid
        const prevCell = ws[XLSX.utils.encode_cell({ r: startRow + i - 1, c: 6 })];
        gValue = prevCell ? prevCell.v : '';
      }
      ws[XLSX.utils.encode_cell({ r: startRow + i, c: 6 })] = { v: gValue, t: typeof gValue === 'number' ? 'n' : 's' };

      // Debug log for column G calculation
      console.log(`Row ${startRow + i + 1}, Column G: ${gValue} (O=${totalQty}, N=${weightPerCarton})`);

      // Net Net Weight (TOTAL N.N.W.) - Col 7
      ws[XLSX.utils.encode_cell({ r: startRow + i, c: 7 })] = { v: row['totalNnw'], t: 'n' };
      // Net Weight (TOTAL N.W.) - Col 8
      ws[XLSX.utils.encode_cell({ r: startRow + i, c: 8 })] = { v: row['totalNw'], t: 'n' };
      // Gross Weight (TOTAL G.W.) - Col 9
      ws[XLSX.utils.encode_cell({ r: startRow + i, c: 9 })] = { v: row['totalGw'], t: 'n' };
      // Carton Size (MEAS. CM) - Col 10
      ws[XLSX.utils.encode_cell({ r: startRow + i, c: 10 })] = { v: row['measCm'], t: 's' };
      // Total CBM (TOTAL CBM) - Col 11
      ws[XLSX.utils.encode_cell({ r: startRow + i, c: 11 })] = { v: row['totalCbm'], t: 'n' };
    });

    // 2. Lower summary box (assume starts at row 22, columns: Colour, OS, ...)
    // Gather color and OS summary
    const colorSummary: Record<string, number> = {};
    orderData.forEach(row => {
      const color = row['color'];
      const os = Number(row['os'] ?? row['carton']); // Use OS if present, else CARTON
      if (!colorSummary[color]) colorSummary[color] = 0;
      colorSummary[color] += os;
    });
    let summaryRow = 22; // 0-indexed (row 23 in Excel)
    Object.entries(colorSummary).forEach(([color, os], idx) => {
      ws[XLSX.utils.encode_cell({ r: summaryRow + idx, c: 0 })] = { v: color, t: 's' }; // Colour
      ws[XLSX.utils.encode_cell({ r: summaryRow + idx, c: 1 })] = { v: os, t: 'n' };    // OS
    });

    // 3. Summary section (assume fixed cells)
    // Total Carton: count of unique cartons
    const totalCarton = new Set(orderData.map(row => row['caseNos'])).size;
    // Total Net Net Weight (kg): sum of TOTAL N.N.W.
    const totalNetNetWeight = orderData.reduce((sum, row) => sum + (Number(row['totalNnw']) || 0), 0);
    // Total Net Weight (kg): sum of TOTAL N.W.
    const totalNetWeight = orderData.reduce((sum, row) => sum + (Number(row['totalNw']) || 0), 0);
    // Total Gross Weight (kg): sum of TOTAL G.W.
    const totalGrossWeight = orderData.reduce((sum, row) => sum + (Number(row['totalGw']) || 0), 0);
    // Total CBM: sum of TOTAL CBM
    const totalCbm = orderData.reduce((sum, row) => sum + (Number(row['totalCbm']) || 0), 0);
    // Write to assumed summary cells (adjust as needed)
    ws[XLSX.utils.encode_cell({ r: 30, c: 1 })] = { v: totalCarton, t: 'n' }; // Total Carton
    ws[XLSX.utils.encode_cell({ r: 31, c: 1 })] = { v: totalNetNetWeight, t: 'n' }; // Total Net Net Weight
    ws[XLSX.utils.encode_cell({ r: 32, c: 1 })] = { v: totalNetWeight, t: 'n' }; // Total Net Weight
    ws[XLSX.utils.encode_cell({ r: 33, c: 1 })] = { v: totalGrossWeight, t: 'n' }; // Total Gross Weight
    ws[XLSX.utils.encode_cell({ r: 34, c: 1 })] = { v: totalCbm, t: 'n' }; // Total CBM
  }
}