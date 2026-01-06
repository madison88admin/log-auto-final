import * as XLSX from 'xlsx';
import { ProcessingResult, ValidationIssue, OrderReport, PackingListRow, ColorSummary, TKListData, ValidationMismatch } from '../types';
import { ExcelReportGenerator } from './ExcelReportGenerator';

// Add Levenshtein distance for fuzzy header matching
function levenshtein(a: string, b: string): number {
  const matrix = Array.from({ length: a.length + 1 }, (_, i) =>
    Array.from({ length: b.length + 1 }, (_, j) => (i === 0 ? j : j === 0 ? i : 0))
  );
  for (let i = 1; i <= a.length; i++) {
    for (let j = 1; j <= b.length; j++) {
      if (a[i - 1] === b[j - 1]) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1, // substitution
          matrix[i][j - 1] + 1,     // insertion
          matrix[i - 1][j] + 1      // deletion
        );
      }
    }
  }
  return matrix[a.length][b.length];
}

export class PackingListProcessor {
  private reportGenerator = new ExcelReportGenerator();

  async processPackingList(file: File): Promise<ProcessingResult> {
    try {
      // Read the Excel file
      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });

      // Check if "PK (2)" sheet exists
      // Find the PK sheet with more flexible matching
      let targetSheet = null;
      let sheetName = '';

      // Try exact match first
      if (workbook.SheetNames.includes('PK (2)')) {
        sheetName = 'PK (2)';
        targetSheet = workbook.Sheets[sheetName];
      } else {
        // Try to find a sheet that contains "PK" and "(2)"
        const pkSheet = workbook.SheetNames.find(name =>
          name.includes('PK') && name.includes('(2)')
        );

        if (pkSheet) {
          sheetName = pkSheet;
          targetSheet = workbook.Sheets[pkSheet];
        } else {
          // Try to find any sheet that starts with "PK"
          const pkSheetAlt = workbook.SheetNames.find(name =>
            name.trim().toLowerCase().startsWith('pk')
          );

          if (pkSheetAlt) {
            sheetName = pkSheetAlt;
            targetSheet = workbook.Sheets[pkSheetAlt];
          }
        }
      }

      if (!targetSheet) {
        throw new Error(`No suitable PK sheet found in the Excel file. Available sheets: ${workbook.SheetNames.join(', ')}. Please ensure there is a sheet named "PK (2)" or similar.`);
      }

      const worksheet = targetSheet;

      // Extract full TK List data
      const tkListData = this.extractTKListData(worksheet);

      if (tkListData.rows.length === 0) {
        throw new Error('No data found in the PK (2) sheet after the header row');
      }

      // Extract and validate data (keeping existing logic for backward compatibility)
      const { validData, validationLog } = this.extractAndValidateData(tkListData.rows);

      if (validData.length === 0) {
        throw new Error('No valid data found after validation');
      }

      // Group by order numbers and generate reports
      const orderReports = this.generateOrderReports(validData);

      // Generate Excel report with full data for validation
      const excelBuffer = await this.reportGenerator.generateExcelReport(orderReports, tkListData, file.name);

      // Perform strict validation
      const strictValidationResults = this.performStrictValidation(orderReports, tkListData);

      return {
        success: true,
        orderReports,
        excelBuffer,
        validationLog,
        strictValidationResults,
        originalFileName: file.name,
        processedAt: new Date()
      };

    } catch (error) {
      throw new Error(`Failed to process packing list: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }

  // Canonical mapping for your provided headers
  private static HEADER_MAPPING: Record<string, string> = {
    'CASE NOS': 'caseNos',
    'SA4 PO NO#': 'sa4PoNo',
    'CARTON PO NO#': 'cartonPoNo',
    'SAP STYLE NO': 'sapStyleNo',
    'STYLE NAME #': 'styleName',
    'S4 Material': 's4Material',
    'Material No#': 'materialNo',
    'COLOR': 'color',
    'Size': 'size',
    'Total QTY': 'totalQty',
    'CARTON': 'carton',
    'QTY / CARTON': 'qtyPerCarton',
    'TOTAL QTY': 'totalQty2',
    'N.N.W / ctn.': 'nnwPerCtn',
    'TOTAL N.N.W.': 'totalNnw',
    'N.W / ctn.': 'nwPerCtn',
    'TOTAL N.W.': 'totalNw',
    'G.W. / ctn': 'gwPerCtn',
    'TOTAL G.W.': 'totalGw',
    'MEAS. CM': 'measCm',
    'TOTAL CBM': 'totalCbm',
  };

  // Helper to normalize header strings (case-insensitive, remove spaces/punctuation)
  private static normalizeHeader(header: string): string {
    return header.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
  }

  private extractTKListData(worksheet: XLSX.WorkSheet): TKListData {
    // Dynamically find the header row (look for 'Carton#', 'Color', etc.)
    const allRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    let headerRowIdx = -1;
    let headerRow: string[] = [];
    let useExactMapping = false;

    for (let i = 0; i < allRows.length; i++) {
      const row = (allRows[i] as (string | undefined)[]).map((cell: any) => (cell || '').toString().trim());
      // Normalize both row and mapping keys for comparison
      const normalizedRow = row.map(PackingListProcessor.normalizeHeader);
      const mappingKeys = Object.keys(PackingListProcessor.HEADER_MAPPING).map(PackingListProcessor.normalizeHeader);
      // If all mapping keys are present in the row, use exact mapping
      const allPresent = mappingKeys.every(key => normalizedRow.includes(key));
      if (allPresent) {
        headerRowIdx = i;
        headerRow = row;
        useExactMapping = true;
        break;
      }
      // Otherwise, use flexible detection as before
      const hasCarton = row.some(cell => {
        const cellLower = cell.toLowerCase();
        return cellLower.includes('carton') ||
          cellLower.includes('box') ||
          cellLower.includes('ctn') ||
          cellLower.includes('carton #') ||
          cellLower.includes('carton#') ||
          cellLower.includes('box #') ||
          cellLower.includes('box#');
      });
      const hasColor = row.some(cell => {
        const cellLower = cell.toLowerCase();
        return cellLower.includes('color') ||
          cellLower.includes('colour') ||
          cellLower.includes('clr') ||
          cellLower.includes('color #') ||
          cellLower.includes('color#') ||
          cellLower.includes('colour #') ||
          cellLower.includes('colour#') ||
          cellLower.includes('clr #') ||
          cellLower.includes('clr#');
      });
      const hasQuantity = row.some(cell => {
        const cellLower = cell.toLowerCase();
        return cellLower.includes('qty') ||
          cellLower.includes('quantity') ||
          cellLower.includes('amount') ||
          cellLower.includes('total') ||
          cellLower.includes('qty #') ||
          cellLower.includes('qty#') ||
          cellLower.includes('quantity #') ||
          cellLower.includes('quantity#');
      });
      const indicators = [hasCarton, hasColor, hasQuantity].filter(Boolean);
      if (indicators.length >= 2 && headerRowIdx === -1) {
        headerRowIdx = i;
        headerRow = row;
        // Don't set useExactMapping
        break;
      }
    }
    if (headerRowIdx === -1) {
      throw new Error('Could not find data table header row. Looking for columns containing "Carton", "Color", "Quantity" or similar terms (case-insensitive).');
    }
    // Extract all data rows below the header
    const dataRows = allRows.slice(headerRowIdx + 1).filter((row: any) =>
      Array.isArray(row) && row.some((cell: any) => cell !== undefined && cell !== null && cell !== '')
    );
    // Compose as array of objects for easier mapping
    // --- PATCH: Fuzzy header mapping and debug output ---
    const requiredFields = Object.values(PackingListProcessor.HEADER_MAPPING);
    console.log('Detected header row:', headerRow);
    const rows: Record<string, any>[] = dataRows.map((row, rowIdx) => {
      const arr = row as any[];
      const obj: Record<string, any> = {};
      headerRow.forEach((col, idx) => {
        let mappedKey: string = col;
        if (useExactMapping) {
          // Use mapping dictionary
          const normCol = PackingListProcessor.normalizeHeader(col);
          mappedKey = Object.entries(PackingListProcessor.HEADER_MAPPING).find(([k]) => PackingListProcessor.normalizeHeader(k) === normCol)?.[1] || col;
        } else {
          // Fuzzy match to mapping keys
          const normCol = PackingListProcessor.normalizeHeader(col);
          const bestMatch = Object.entries(PackingListProcessor.HEADER_MAPPING)
            .map(([k, v]) => ({ k, v, dist: levenshtein(PackingListProcessor.normalizeHeader(k), normCol) }))
            .sort((a, b) => a.dist - b.dist)[0];
          if (bestMatch && bestMatch.dist <= 2) {
            mappedKey = bestMatch.v;
          } else {
            mappedKey = col;
          }
        }
        obj[mappedKey] = arr[idx];
      });
      // Debug: log the mapped row
      console.log(`Row ${rowIdx + 1} mapped:`, obj);
      // Warn if any required fields are missing
      const missing = requiredFields.filter(f => !(f in obj));
      if (missing.length > 0) {
        console.warn(`Row ${rowIdx + 1} is missing fields:`, missing);
      }
      return obj;
    });
    // --- END PATCH ---
    return {
      headers: headerRow,
      rows,
      orderGroups: this.groupRowsByOrder(rows)
    };
  }

  private findOrderNumber(row: Record<string, any>): string {
    // More flexible order number detection with case-insensitive matching
    const orderKeys = [
      'Order Number', 'Order', 'PO', 'PO Number', 'OrderNumber', 'Order_Number',
      'Purchase Order', 'PurchaseOrder', 'PO#', 'Order#', 'Order No', 'OrderNo',
      'PO-Line', 'PO Line', 'POLine', 'Order Line', 'OrderLine',
      // Add variations with different capitalization
      'ORDER NUMBER', 'ORDER', 'PO NUMBER', 'ORDER NUMBER',
      'PURCHASE ORDER', 'PURCHASEORDER', 'ORDER NO', 'ORDERNO',
      'PO-LINE', 'PO LINE', 'ORDER LINE', 'ORDERLINE'
    ];

    for (const key of orderKeys) {
      if (row[key] !== undefined && row[key] !== null && row[key] !== '') {
        return String(row[key]).trim();
      }
    }

    // If no exact match, try partial matches (case-insensitive)
    for (const [key, value] of Object.entries(row)) {
      if (value !== undefined && value !== null && value !== '') {
        const keyLower = key.toLowerCase();
        if (keyLower.includes('order') || keyLower.includes('po') || keyLower.includes('purchase')) {
          return String(value).trim();
        }
      }
    }

    return 'Unknown';
  }

  private performStrictValidation(orderReports: OrderReport[], tkListData: TKListData): ValidationMismatch[] {
    const mismatches: ValidationMismatch[] = [];

    // For each order report, compare with original TK List data
    orderReports.forEach(orderReport => {
      const originalOrderData = tkListData.orderGroups[orderReport.orderNumber] || [];

      // Compare each row in the generated report with the original data
      // This is a simplified comparison - in practice, you'd need to map the exact cells
      originalOrderData.forEach((originalRow, rowIndex) => {
        Object.keys(originalRow).forEach(fieldName => {
          const originalValue = originalRow[fieldName];
          // Find corresponding value in generated report
          // This would need to be implemented based on the actual mapping logic
          const generatedValue = this.findGeneratedValue(orderReport, fieldName, rowIndex);

          if (originalValue !== generatedValue) {
            mismatches.push({
              orderNumber: orderReport.orderNumber,
              row: rowIndex + 1,
              column: fieldName,
              originalValue,
              generatedValue,
              fieldName
            });
          }
        });
      });
    });

    return mismatches;
  }

  private findGeneratedValue(orderReport: OrderReport, fieldName: string, rowIndex: number): any {
    // This is a placeholder - implement based on actual mapping logic
    // For now, return a dummy value to demonstrate the structure
    return `Generated_${fieldName}_${rowIndex}`;
  }

  private extractAndValidateData(jsonData: Record<string, any>[]): { validData: PackingListRow[], validationLog: ValidationIssue[] } {
    const validData: PackingListRow[] = [];
    const validationLog: ValidationIssue[] = [];

    jsonData.forEach((row, index) => {
      // Use your mapped keys here
      const color = row['color'];
      const cartonPoNo = row['cartonPoNo'];
      const totalQty = row['totalQty'] ?? row['totalQty2']; // handle both if needed

      // Skip completely empty rows
      if (!color && !cartonPoNo && !totalQty) {
        return;
      }

      // Add your validation logic here, or just push all non-empty rows as valid
      validData.push(row as PackingListRow);
    });

    return { validData, validationLog };
  }

  private generateOrderReports(validData: PackingListRow[], modelNamesByOrder?: Record<string, string>): OrderReport[] {
    // Group data by order number
    const orderGroups = validData.reduce((groups, row) => {
      // Default orderNumber to 'Unknown' if missing or falsy
      const orderNumber = row.orderNumber ? String(row.orderNumber) : 'Unknown';
      if (!groups[orderNumber]) {
        groups[orderNumber] = [];
      }
      groups[orderNumber].push(row);
      return groups;
    }, {} as Record<string, PackingListRow[]>);

    // Generate reports for each order
    return Object.entries(orderGroups).map(([orderNumber, data]) => {
      const colorSummary = this.generateColorSummary(data);
      // Guard: ensure color.quantity is always a number
      const totalQuantity = colorSummary.reduce((sum, color) => sum + (typeof color.quantity === 'number' && !isNaN(color.quantity) ? color.quantity : 0), 0);
      return {
        orderNumber,
        data,
        colorSummary,
        totalQuantity,
        modelName: modelNamesByOrder ? modelNamesByOrder[orderNumber] : undefined
      };
    });
  }

  private generateColorSummary(data: PackingListRow[]): ColorSummary[] {
    const colorTotals = data.reduce((totals, row) => {
      const color = row.color || 'Unknown';
      // Always parse quantity as number, default to 0 if invalid
      let qty = Number(row.quantity);
      if (!isFinite(qty)) qty = 0;
      totals[color] = (totals[color] || 0) + qty;
      return totals;
    }, {} as Record<string, number>);

    return Object.entries(colorTotals)
      .map(([color, quantity]) => ({ color, quantity }))
      .sort((a, b) => b.quantity - a.quantity); // Sort by quantity descending
  }

  private groupRowsByOrder(rows: Record<string, any>[]): Record<string, Record<string, any>[]> {
    const orderGroups: Record<string, Record<string, any>[]> = {};
    rows.forEach(row => {
      const orderNumber = this.findOrderNumber(row);
      if (!orderGroups[orderNumber]) {
        orderGroups[orderNumber] = [];
      }
      orderGroups[orderNumber].push(row);
    });
    return orderGroups;
  }
}