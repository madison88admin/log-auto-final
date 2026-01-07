import { useState } from 'react';
import { FileUpload } from './components/FileUpload';
import { ProcessingStatus } from './components/ProcessingStatus';
import { ReportResults } from './components/ReportResults';
import { ValidationResults } from './components/ValidationResults';
import { StrictValidationResults } from './components/StrictValidationResults';
import { ExcelHandsontablePreview } from './components/ExcelHandsontablePreview';
import { Header } from './components/Header';
import { ProcessingResult } from './types';
import ExcelJS from 'exceljs';
import LuckysheetPreview from './components/LuckysheetPreview';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';

function App() {
  const [uploadedFile, setUploadedFile] = useState<File | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [processingResult, setProcessingResult] = useState<ProcessingResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [generatedExcelBuffer, setGeneratedExcelBuffer] = useState<ArrayBuffer | null>(null);
  const [parsedTables, setParsedTables] = useState<{ table: any[][], modelName: string, mergedCellText?: string }[]>([]); const [selectedTableIdx, setSelectedTableIdx] = useState(0);
  // Add state to store all generated report buffers
  const [reportBuffers, setReportBuffers] = useState<ArrayBuffer[]>([]);
  // Add state to store sheet names for download
  const [reportSheetNames, setReportSheetNames] = useState<string[]>([]);
  // Add state for merged cell preview
  const [mergedCellText, setMergedCellText] = useState<string>('');

  const handleFileUpload = (file: File) => {
    setUploadedFile(file);
    setProcessingResult(null);
    setGeneratedExcelBuffer(null);
    setError(null);
    setParsedTables([]);
    setSelectedTableIdx(0);
    setMergedCellText('');
  };

  // Helper to generate all report buffers and sheet names
  async function generateAllReportBuffers(parsedTables: { table: any[][], modelName: string }[], _uploadedFile: File | null) {
    let sheetNames: string[] = [];
    const buffers: ArrayBuffer[] = [];
    for (let t = 0; t < parsedTables.length; t++) {
      const { table } = parsedTables[t];
      const [headerRow, ...dataRows] = table;
      const idxPoNo = headerRow.indexOf('SA4 PO NO#');
      let poNo = idxPoNo !== -1 ? String(dataRows[0]?.[idxPoNo] || `Report ${t + 1}`) : `Report ${t + 1}`;
      let sheetName = poNo.replace(/[\\/?*\[\]:]/g, '_').substring(0, 31);
      let origSheetName = sheetName;
      let suffix = 2;
      while (sheetNames.includes(sheetName)) {
        sheetName = origSheetName.substring(0, 28) + '_' + suffix;
        suffix++;
      }
      sheetNames.push(sheetName);
      // Load template for each report
      const response = await fetch('/ReportTemplate.xlsx');
      const arrayBuffer = await response.arrayBuffer();
      const templateWb = new ExcelJS.Workbook();
      await templateWb.xlsx.load(arrayBuffer);
      const ws = templateWb.getWorksheet('Report');
      if (!ws) throw new Error("Worksheet 'Report' not found in template.");
      // Fill the worksheet as before (reuse fillData logic)
      await (async () => {
        // Debug: log the parsed table structure
        console.log('headerRow:', headerRow);
        console.log('dataRows:', dataRows);

        // Column indices
        const idxColor = headerRow.findIndex(h => h && h.toString().replace(/\s+/g, '').toUpperCase().includes('COLOR'));
        const idxCaseNos = headerRow.indexOf('CASE NOS');
        const idxS4Material = headerRow.indexOf('S4 Material');
        const idxECCMaterial = headerRow.indexOf('Material No#');
        const sizeNames = ['OS', 'XS', 'S', 'M', 'L', 'XL', 'XXL'];

        // Helper to safely get a value
        const safe = (row: any[], idx: number) => (idx >= 0 && row && row[idx] !== undefined ? row[idx] : '');

        // Model Name logic: find the first CASE NOS cell, get 2 rows above
        let modelName = '';
        for (let i = 0; i < table.length; i++) {
          if (safe(table[i], idxCaseNos) && String(safe(table[i], idxCaseNos)).toLowerCase().includes('case')) {
            modelName = safe(table[i - 2], idxCaseNos);
            break;
          }
        }
        // Prefer parsedTables[t].modelName if available
        if (parsedTables[t]?.modelName) {
          modelName = parsedTables[t].modelName;
        }
        // Get S4 HANA SKU and trim last 2 digits
        const firstS4HanaSKU = safe(dataRows[0], idxS4Material)?.toString() || '';
        const poLineValue = firstS4HanaSKU.length > 2 ? firstS4HanaSKU.slice(0, -2) : firstS4HanaSKU;
        // Set PO-Line (D14) and Model # (D16) to trimmed S4 HANA SKU
        ws.getCell('D14').value = poLineValue;
        ws.getCell('D16').value = poLineValue;
        // Set SAP PO# (E14) to sheet name
        ws.getCell('E14').value = sheetName;
        // Model Name (E16) remains as before
        ws.getCell('E16').value = modelName;
        ws.getCell('E7').value = '';

        // --- UPDATE TEMPLATE HEADERS TO INCLUDE NEW N.N.W. COLUMN ---
        // The template header row is at row 19
        // Shift all headers from P onwards to the right, then insert new N.N.W. header
        ws.getCell('P19').value = 'TOTAL\nN.N.W.';
        ws.getCell('P19').alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        ws.getCell('Q19').value = 'Net Weight';
        ws.getCell('Q19').alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        ws.getCell('R19').value = 'Gross Weight';
        ws.getCell('R19').alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

        // Carton Size spans S19:U19 (merged)
        // First unmerge if already merged to avoid conflicts
        try {
          ws.unMergeCells('S19:U19');
        } catch (e) {
          // Ignore if not merged
        }
        // Also check for the old template merge (R19:T19) and unmerge it
        try {
          ws.unMergeCells('R19:T19');
        } catch (e) {
          // Ignore if not merged
        }
        ws.mergeCells('S19:U19');
        ws.getCell('S19').value = 'Carton Size';
        ws.getCell('S19').alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

        ws.getCell('V19').value = 'CBM';
        ws.getCell('V19').alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        ws.getCell('W19').value = 'Total CBM';
        ws.getCell('W19').alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

        // --- Extend "Required Shipping Data" yellow header to include Total CBM (N18:W18) ---
        try {
          ws.unMergeCells('N18:V18'); // Unmerge old range if exists
        } catch (e) {
          // Ignore if not merged
        }
        ws.mergeCells('N18:W18'); // Merge to include W column
        ws.getCell('N18').value = 'Required Shipping Data';
        ws.getCell('N18').fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFFF00' } // Yellow
        };
        ws.getCell('N18').font = { bold: true };
        ws.getCell('N18').alignment = { horizontal: 'center', vertical: 'middle' };

        // --- DYNAMIC MAIN TABLE ROWS AT C20 ---
        const mainTableStart = 20; // C20 is row 20
        const numDataRows = dataRows.length;
        ws.insertRows(mainTableStart, Array(numDataRows).fill([]));

        // 1. Propagate carton numbers for all rows
        let lastCartonNo = '';
        const effectiveCartonNos = dataRows.map(row => {
          const cartonNo = safe(row, idxCaseNos)?.toString().trim();
          if (cartonNo) {
            lastCartonNo = cartonNo;
            return cartonNo;
          }
          return lastCartonNo;
        });

        // 2. Build cartonCountMap using propagated carton numbers
        const cartonCountMap: Record<string, number> = {};
        effectiveCartonNos.forEach(cartonNo => {
          if (!cartonNo) return;
          cartonCountMap[cartonNo] = (cartonCountMap[cartonNo] || 0) + 1;
        });

        // 3. Write rows using propagated carton number for split-carton logic
        // For merging: track start/end row for each carton group
        const cartonRowRanges: Record<string, { start: number, end: number }> = {};
        effectiveCartonNos.forEach((cartonNo, i) => {
          const rowNum = mainTableStart + i;
          if (!(cartonNo in cartonRowRanges)) {
            cartonRowRanges[cartonNo] = { start: rowNum, end: rowNum };
          } else {
            cartonRowRanges[cartonNo].end = rowNum;
          }
        });

        // Copy formatting for new rows from the template's main table row (row 20 before insertion)
        const styleRow = ws.getRow(mainTableStart + numDataRows); // This is the original template row 20
        for (let i = 0; i < numDataRows; i++) {
          const newRow = ws.getRow(mainTableStart + i);
          for (let j = 1; j <= ws.columnCount; j++) {
            const styleCell = styleRow.getCell(j);
            const newCell = newRow.getCell(j);
            newCell.style = { ...styleCell.style };
            if (styleCell.numFmt) newCell.numFmt = styleCell.numFmt;
            if (styleCell.alignment) newCell.alignment = styleCell.alignment;
            if (styleCell.border) newCell.border = styleCell.border;
            if (styleCell.fill) newCell.fill = styleCell.fill;
            // 1. Formula assignment
            if (styleCell.formula) {
              newCell.value = { formula: styleCell.formula, result: undefined };
            }
          }
        }

        let prevColor = '';
        let prevS4SKU = '';
        let prevECC = '';

        console.log('Header Row:', headerRow);
        function findHeaderIndex(headerRow: any[], search: string) {
          return headerRow.findIndex((h: any) =>
            h && h.toString().replace(/[^a-zA-Z0-9]/g, '').toLowerCase().includes(search)
          );
        }
        const idxUnitsCrt = findHeaderIndex(headerRow, 'qtycarton');
        const idxCarton = findHeaderIndex(headerRow, 'carton');
        const idxMeasCm = findHeaderIndex(headerRow, 'meascm');
        console.log('idxUnitsCrt:', idxUnitsCrt, 'idxCarton:', idxCarton, 'idxMeasCm:', idxMeasCm);

        dataRows.forEach((row, i) => {
          const rowNum = mainTableStart + i;
          // 4. Only write Carton# for the first row in the group, else leave blank
          const cartonNo = effectiveCartonNos[i];
          if (cartonRowRanges[cartonNo].start === rowNum) {
            ws.getCell(`C${rowNum}`).value = cartonNo;
          } else {
            ws.getCell(`C${rowNum}`).value = '';
          }

          // Color (D)
          let color = safe(row, idxColor);
          if (!color) color = prevColor;
          ws.getCell(`D${rowNum}`).value = color;
          prevColor = color;

          // S4 HANA SKU (E)
          let s4sku = safe(row, idxS4Material);
          if (!s4sku) s4sku = prevS4SKU;
          ws.getCell(`E${rowNum}`).value = s4sku;
          prevS4SKU = s4sku;

          // ECC Material No (F)
          let ecc = safe(row, idxECCMaterial);
          if (!ecc) ecc = prevECC;
          ws.getCell(`F${rowNum}`).value = typeof ecc === 'string' && (ecc.startsWith('X') || ecc.startsWith('L')) ? ecc : prevECC;
          prevECC = typeof ws.getCell(`F${rowNum}`).value === 'string' ? ws.getCell(`F${rowNum}`).value as string : '';

          // --- SPLIT CARTON LOGIC FOR OS COLUMN ---
          // For the OS column in the report:
          //   - If split carton (carton number appears more than once), use value from Column J (index 9, 0-based)
          //   - If single carton, use value from Column L (index 11, 0-based)
          let osValue;
          if (cartonNo && cartonCountMap[cartonNo] > 1) {
            // Split carton: use Column J (index 9)
            osValue = safe(row, 9);
          } else {
            // Single-color carton: use Column L (index 11)
            osValue = safe(row, 11);
          }
          ws.getCell(`G${rowNum}`).value = osValue;

          // Fill other size columns as before
          sizeNames.forEach((size, j) => {
            const idx = headerRow.indexOf(size);
            if (idx !== -1) {
              // Only fill H, I, J, K, L, M, N (columns 8-14) if not OS (G)
              if (j === 0) return; // skip OS, already filled
              ws.getCell(String.fromCharCode(71 + j) + rowNum).value = safe(row, idx);
            }
          });

          // Map report columns to uploaded file columns by letter
          // Input file structure: K=10(CARTON), L=11, M=12(TOTAL QTY), N=13(N.N.W/ctn), O=14(TOTAL N.N.W.), P=15(N.W/ctn), Q=16(TOTAL N.W.), R=17(G.W./ctn), S=18(TOTAL G.W.), T=19, U=20, V=21
          // Output columns in report template
          ws.getCell(`N${rowNum}`).value = safe(row, 10); // N ← K (CARTON / Units/CRT)
          ws.getCell(`O${rowNum}`).value = safe(row, 12); // O ← M (TOTAL QTY / Total Unit)
          // Set values with 3 decimal format
          const pCell = ws.getCell(`P${rowNum}`);
          pCell.value = parseFloat(safe(row, 14)); // P ← O (TOTAL N.N.W.)
          pCell.numFmt = '0.000';

          const qCell = ws.getCell(`Q${rowNum}`);
          qCell.value = parseFloat(safe(row, 16)); // Q ← Q (TOTAL N.W. / Net Weight)
          qCell.numFmt = '0.000';

          const rCell = ws.getCell(`R${rowNum}`);
          rCell.value = parseFloat(safe(row, 18)); // R ← S (TOTAL G.W. / Gross Weight)
          rCell.numFmt = '0.000';

          ws.getCell(`S${rowNum}`).value = safe(row, 19); // S ← T (Length)
          ws.getCell(`T${rowNum}`).value = safe(row, 20); // T ← U (Width)
          ws.getCell(`U${rowNum}`).value = safe(row, 21); // U ← V (Height)

          const vCell = ws.getCell(`V${rowNum}`);
          vCell.value = parseFloat(safe(row, 22)); // V ← W (CBM)
          vCell.numFmt = '0.000';

          const wCell = ws.getCell(`W${rowNum}`);
          wCell.value = parseFloat(safe(row, 22)); // W ← W (TOTAL CBM - same as CBM)
          wCell.numFmt = '0.000';
        });

        // 5. Merge Carton# cells for each group in the worksheet
        Object.values(cartonRowRanges).forEach(({ start, end }) => {
          if (end > start) {
            ws.mergeCells(`C${start}:C${end}`);
            // Also merge columns N–W for this group (skip N for carton count)
            const colLetters = ['O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W'];
            colLetters.forEach(col => {
              ws.mergeCells(`${col}${start}:${col}${end}`);
            });
          }
        });



        // --- SUMMARY AND COLOR BREAKDOWN ---
        // Move summary and color breakdown to start 1 row below the last data row
        const summaryStartRow = mainTableStart + numDataRows + 1;
        // Merge D and E for the summary name
        ws.mergeCells(`D${summaryStartRow}:E${summaryStartRow}`);
        ws.getCell(`D${summaryStartRow}`).value = 'Summary';
        ws.getCell(`D${summaryStartRow}`).alignment = { horizontal: 'center', vertical: 'middle' };

        // Write summary titles in D, values in E
        // Calculate summary values from the ORIGINAL dataRows (not from worksheet cells to avoid rounding errors)
        let totalCarton = 0;
        let totalNetNetWeight = 0;
        let totalNetWeight = 0;
        let totalGrossWeight = 0;
        let totalCBM = 0;
        for (let i = 0; i < numDataRows; i++) {
          // Sum from original data
          // Column 10 = Units/CRT (CARTON), columns 14, 16, 18, 22 for weights and CBM
          // Round each value to 3 decimals before summing to avoid floating point precision errors
          totalCarton += Math.round((parseFloat(safe(dataRows[i], 10)) || 0) * 1000) / 1000;
          totalNetNetWeight += Math.round((parseFloat(safe(dataRows[i], 14)) || 0) * 1000) / 1000;
          totalNetWeight += Math.round((parseFloat(safe(dataRows[i], 16)) || 0) * 1000) / 1000;
          totalGrossWeight += Math.round((parseFloat(safe(dataRows[i], 18)) || 0) * 1000) / 1000;
          totalCBM += Math.round((parseFloat(safe(dataRows[i], 22)) || 0) * 1000) / 1000;
        }
        // Round final totals to 3 decimals
        totalCarton = Math.round(totalCarton * 1000) / 1000;
        totalNetNetWeight = Math.round(totalNetNetWeight * 1000) / 1000;
        totalNetWeight = Math.round(totalNetWeight * 1000) / 1000;
        totalGrossWeight = Math.round(totalGrossWeight * 1000) / 1000;
        totalCBM = Math.round(totalCBM * 1000) / 1000;
        const summaryData = [
          { title: 'Total Carton', value: totalCarton, format: '0.000' },
          { title: 'Total Net Net Weight', value: totalNetNetWeight, format: '0.000' },
          { title: 'Total Net Weight', value: totalNetWeight, format: '0.000' },
          { title: 'Total Gross Weight', value: totalGrossWeight, format: '0.000' },
          { title: 'Total CBM', value: totalCBM, format: '0.000' }
        ];
        summaryData.forEach((item, i) => {
          ws.getCell(`D${summaryStartRow + 1 + i}`).value = item.title;
          const valueCell = ws.getCell(`E${summaryStartRow + 1 + i}`);
          valueCell.value = item.value;
          if (item.format) {
            valueCell.numFmt = item.format;
          }
        });

        // --- COLOR BREAKDOWN FROM GENERATED WORKSHEET ---
        // Columns H:N are OS, XS, S, M, L, XL, XXL (col 8-14, 1-based)
        const colorMap: Record<string, number[]> = {};
        // Map each size to its worksheet column letter (adjusted for your template)
        const sizeColLetters = ['G', 'H', 'I', 'J', 'K', 'L', 'M']; // OS, XS, S, M, L, XL, XXL
        for (let i = 0; i < numDataRows; i++) {
          const rowNum = mainTableStart + i;
          const color = ws.getCell(`D${rowNum}`).value?.toString().trim() || '';
          if (!color) continue;
          if (!colorMap[color]) colorMap[color] = [0, 0, 0, 0, 0, 0, 0];
          const cartonNo = effectiveCartonNos[i];
          // OS column (index 0): use split carton logic
          if (cartonNo && cartonCountMap[cartonNo] > 1) {
            // Split carton: use OS (G)
            colorMap[color][0] += parseFloat(ws.getCell(`G${rowNum}`).value as string) || 0;
          } else {
            // Single carton: use Total Unit (O)
            colorMap[color][0] += parseFloat(ws.getCell(`O${rowNum}`).value as string) || 0;
          }
          // Other sizes: use previous logic (from worksheet columns)
          for (let j = 1; j < 7; j++) {
            const cellVal = parseFloat(ws.getCell(`${sizeColLetters[j]}${rowNum}`).value as string) || 0;
            colorMap[color][j] += cellVal;
          }
        }
        // --- DYNAMIC COLOR BREAKDOWN TABLE ---
        // Place color breakdown headers adjacent to the Total Carton row in the summary
        const colorBreakdownStartRow = summaryStartRow + 1; // This is the Total Carton row
        const colorBreakdownStartCol = 6; // Column F
        const colorHeaders = ['Color', ...sizeNames, 'Total'];
        const borderStyle = { style: 'thin' as ExcelJS.BorderStyle };
        const headerBorder = { top: borderStyle, left: borderStyle, bottom: borderStyle, right: borderStyle };

        // Write headers with borders
        colorHeaders.forEach((header, i) => {
          const cell = ws.getCell(colorBreakdownStartRow, colorBreakdownStartCol + i);
          cell.value = header;
          cell.style = { font: { bold: true }, alignment: { horizontal: 'center', vertical: 'middle' } };
          cell.border = headerBorder;
        });

        // Filter out empty or falsy color keys
        const validColors = Object.keys(colorMap).filter(c => c && c.trim() !== '');

        // Write color breakdown rows dynamically with borders
        let colorIdx = 0;
        for (const color of validColors) {
          ws.getCell(colorBreakdownStartRow + 1 + colorIdx, colorBreakdownStartCol).value = color;
          ws.getCell(colorBreakdownStartRow + 1 + colorIdx, colorBreakdownStartCol).style = { alignment: { horizontal: 'left' } };
          ws.getCell(colorBreakdownStartRow + 1 + colorIdx, colorBreakdownStartCol).border = headerBorder;
          for (let i = 0; i < 7; i++) {
            const cell = ws.getCell(colorBreakdownStartRow + 1 + colorIdx, colorBreakdownStartCol + 1 + i);
            cell.value = colorMap[color][i];
            cell.style = { alignment: { horizontal: 'center' } };
            cell.border = headerBorder;
          }
          // Total for the color
          const totalCell = ws.getCell(colorBreakdownStartRow + 1 + colorIdx, colorBreakdownStartCol + 8);
          totalCell.value = colorMap[color].reduce((a, b) => a + b, 0);
          totalCell.style = { font: { bold: true }, alignment: { horizontal: 'center' } };
          totalCell.border = headerBorder;
          colorIdx++;
        }
        // Write the total row for color breakdown with borders
        ws.getCell(colorBreakdownStartRow + 1 + colorIdx, colorBreakdownStartCol).value = 'Total';
        ws.getCell(colorBreakdownStartRow + 1 + colorIdx, colorBreakdownStartCol).style = { font: { bold: true }, alignment: { horizontal: 'left' } };
        ws.getCell(colorBreakdownStartRow + 1 + colorIdx, colorBreakdownStartCol).border = headerBorder;
        for (let i = 0; i < 7; i++) {
          const cell = ws.getCell(colorBreakdownStartRow + 1 + colorIdx, colorBreakdownStartCol + 1 + i);
          cell.value = Object.values(colorMap).reduce((acc, arr) => acc + arr[i], 0);
          cell.style = { font: { bold: true }, alignment: { horizontal: 'center' } };
          cell.border = headerBorder;
        }
        const grandTotalCell = ws.getCell(colorBreakdownStartRow + 1 + colorIdx, colorBreakdownStartCol + 8);
        grandTotalCell.value = Object.values(colorMap).reduce((acc, arr) => acc + arr.reduce((a, b) => a + b, 0), 0);
        grandTotalCell.style = { font: { bold: true }, alignment: { horizontal: 'center' } };
        grandTotalCell.border = headerBorder;

        // --- REMOVE ALL IMAGES FROM WORKSHEET ---
        try {
          // Remove all images if they exist
          if (typeof ws.getImages === 'function') {
            const images = ws.getImages();
            if (images && images.length > 0) {
              // Clear all images by setting model properties to empty arrays
              if (ws.model && (ws.model as any).media) {
                (ws.model as any).media = [];
              }
            }
          }
          // Also try to clear via direct property access
          if ((ws as any)._media) {
            (ws as any)._media = [];
          }
          if ((ws as any)._images) {
            (ws as any)._images = [];
          }
        } catch (error) {
          console.error('Error removing images:', error);
          // Continue even if image removal fails
        }
      })();
      // Export this workbook to a buffer
      const buffer = await templateWb.xlsx.writeBuffer();
      buffers.push(buffer);
    }
    return { buffers, sheetNames };
  }

  const handleGenerateReports = async () => {
    if (!parsedTables.length) return;
    setIsProcessing(true);
    setError(null);
    try {
      // Generate all report buffers and sheet names
      const { buffers, sheetNames } = await generateAllReportBuffers(parsedTables, uploadedFile);
      setReportBuffers(buffers);
      setReportSheetNames(sheetNames);
      // For now, just use the first buffer for preview
      setGeneratedExcelBuffer(buffers[0] || null);
      setProcessingResult({
        success: true,
        orderReports: [],
        excelBuffer: buffers[0] || null,
        validationLog: [],
        strictValidationResults: [],
        originalFileName: `${(uploadedFile?.name || 'Report').replace(/\.xlsx?$/i, '')}-Report.xlsx`,
        processedAt: new Date()
      });
    } catch (err) {
      setError(err instanceof Error ? err.message : 'An error occurred while processing the file');
    } finally {
      setIsProcessing(false);
    }
  };

  // Add a function to download a single report
  const handleDownloadSingleReport = (idx: number) => {
    if (!reportBuffers[idx] || !reportSheetNames[idx]) return;
    const baseName = (uploadedFile?.name || 'Report').replace(/\.xlsx?$/i, '');
    const fileName = `${baseName}-${reportSheetNames[idx]}.xlsx`;
    saveAs(new Blob([reportBuffers[idx]], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), fileName);
  };

  const handleReset = () => {
    setUploadedFile(null);
    setProcessingResult(null);
    setGeneratedExcelBuffer(null);
    setError(null);
    setIsProcessing(false);
    setParsedTables([]);
    setSelectedTableIdx(0);
  };

  // Add the handler function if not present
  const handleExportAllAsSingleExcel = async () => {
    if (!uploadedFile) return;
    setIsProcessing(true);
    setError(null);
    try {
      const formData = new FormData();
      formData.append('file', uploadedFile);

      const response = await fetch('http://localhost:8000/generate-reports/', {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) throw new Error('Failed to generate combined report');

      const blob = await response.blob();
      // Use uploaded file name + 'Report.xlsx' for the download
      const baseName = (uploadedFile.name || 'Report').replace(/\.xlsx?$/i, '');
      saveAs(blob, `${baseName}Report.xlsx`);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'An error occurred while exporting the combined report');
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="min-h-screen bg-gray-50">
      <Header />
      <main className="w-full px-2 py-8">
        <div className="space-y-8">
          {/* File Upload Section */}
          <section className="card">
            <h2 className="text-xl font-semibold text-gray-900 mb-4">
              Upload TK List
            </h2>
            <FileUpload
              onFileUpload={handleFileUpload}
              uploadedFile={uploadedFile}
              disabled={isProcessing}
            />
          </section>

          {/* Processing Controls */}
          {uploadedFile && !processingResult && (
            <section className="card">
              <div className="flex items-center justify-between">
                <div>
                  <h3 className="text-lg font-medium text-gray-900">
                    Ready to Process
                  </h3>
                  <p className="text-gray-600 mt-1">
                    File: {uploadedFile.name}
                  </p>
                </div>
                <div className="flex gap-3">
                  <button
                    onClick={handleReset}
                    className="btn-secondary"
                    disabled={isProcessing}
                  >
                    Reset
                  </button>
                  <button
                    onClick={handleGenerateReports}
                    className="btn-primary"
                    disabled={isProcessing}
                  >
                    Generate Reports
                  </button>
                </div>
              </div>
            </section>
          )}

          {/* Side-by-side Excel Previews (Excel-like) */}
          {(uploadedFile || generatedExcelBuffer) && (
            <section className="card" style={{ padding: 0, maxWidth: '100%' }}>
              <div className="flex flex-col md:flex-row gap-6 w-full">
                <div className="flex-1 min-w-0">
                  <ExcelHandsontablePreview
                    excelFile={uploadedFile}
                    title="PK Table"
                    onTablesExtracted={tables => {
                      setParsedTables(tables);
                      setMergedCellText(tables[0]?.mergedCellText || '');
                    }}
                    onSelectedTableChange={(_table, _modelName) => {
                      // Find the index of the selected table
                      const idx = parsedTables.findIndex(t => t.table === _table && t.modelName === _modelName);
                      if (idx >= 0) setSelectedTableIdx(idx);
                    }}
                  />
                </div>
                <div className="flex-1 min-w-0 flex flex-col">
                  {mergedCellText && (
                    <div style={{
                      margin: '0 0 16px 0',
                      padding: '16px 20px',
                      background: 'linear-gradient(90deg, #f8fafc 80%, #e0e7ef 100%)',
                      border: '1.5px solid #cbd5e1',
                      borderRadius: 10,
                      fontFamily: 'Inter, Segoe UI, Arial, sans-serif',
                      whiteSpace: 'pre-line',
                      minHeight: 48,
                      fontSize: 14,
                      color: '#222',
                      alignSelf: 'stretch',
                      width: '100%',
                      boxSizing: 'border-box',
                      boxShadow: '0 2px 8px rgba(0,0,0,0.07)',
                      display: 'flex',
                      alignItems: 'flex-start',
                      gap: 12,
                    }}>
                      <span style={{ display: 'flex', alignItems: 'center', marginRight: 8 }}>
                        <svg width="20" height="20" fill="none" stroke="#2563eb" strokeWidth="2" viewBox="0 0 24 24" style={{ marginRight: 6 }}><rect x="4" y="4" width="16" height="16" rx="4" /><path d="M8 9h8M8 13h5" /></svg>
                        <b style={{ color: '#2563eb', fontSize: 15 }}>Ship to:</b>
                      </span>
                      <span style={{ flex: 1 }}>{mergedCellText}</span>
                    </div>
                  )}
                  {reportBuffers[selectedTableIdx] && (
                    <LuckysheetPreview
                      excelBlob={new Blob([reportBuffers[selectedTableIdx]], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })}
                    />
                  )}
                </div>
              </div>
            </section>
          )}

          {/* Export/Process Buttons - move above validation results */}
          {processingResult && reportBuffers.length > 0 && (
            <div className="flex flex-wrap gap-2 mt-4">
              <button
                onClick={() => handleDownloadSingleReport(selectedTableIdx)}
                className="btn-secondary flex items-center gap-2"
                disabled={
                  !reportBuffers[selectedTableIdx] || !reportSheetNames[selectedTableIdx]
                }
              >
                {/* Download icon */}
                <svg className="w-5 h-5" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" d="M4 16v2a2 2 0 002 2h12a2 2 0 002-2v-2M7 10l5 5m0 0l5-5m-5 5V4" /></svg>
                Export Single Report{reportSheetNames[selectedTableIdx] ? ` - ${reportSheetNames[selectedTableIdx]}` : ''}
              </button>
              <button
                onClick={handleExportAllAsSingleExcel}
                className="btn-primary flex items-center gap-2"
              >
                {/* Excel icon */}
                <svg className="w-5 h-5" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" d="M12 4v16m8-8H4" /></svg>
                Export All
              </button>
              <button
                onClick={handleReset}
                className="btn-secondary flex items-center gap-2"
                style={{ marginLeft: 'auto' }}
              >
                {/* Refresh/plus icon */}
                <svg className="w-5 h-5" fill="none" stroke="currentColor" strokeWidth="2" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" d="M12 4v16m8-8H4" /></svg>
                Process New File
              </button>
            </div>
          )}

          {/* Processing Status */}
          {isProcessing && (
            <ProcessingStatus />
          )}

          {/* Error Display */}
          {error && (
            <section className="card border-red-200 bg-red-50">
              <div className="flex items-start gap-3">
                <div className="w-5 h-5 text-red-500 mt-0.5">
                  <svg fill="currentColor" viewBox="0 0 20 20">
                    <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.707 7.293a1 1 0 00-1.414 1.414L8.586 10l-1.293 1.293a1 1 0 101.414 1.414L10 11.414l1.293 1.293a1 1 0 001.414-1.414L11.414 10l1.293-1.293a1 1 0 00-1.414-1.414L10 8.586 8.707 7.293z" clipRule="evenodd" />
                  </svg>
                </div>
                <div>
                  <h3 className="text-lg font-medium text-red-900">
                    Processing Error
                  </h3>
                  <p className="text-red-700 mt-1">{error}</p>
                </div>
              </div>
            </section>
          )}

          {/* Results */}
          {processingResult && (
            <>
              <ReportResults
                result={processingResult}
                onReset={handleReset}
              />
              <ValidationResults
                validationLog={processingResult.validationLog}
              />
              <StrictValidationResults
                validationResults={processingResult.strictValidationResults}
                onReset={handleReset}
              />
            </>
          )}
        </div>
      </main>
    </div>
  );
}

export default App;

export async function generateExcelReportWithExcelJS(
  templateUrl: string,
  fillData: (workbook: ExcelJS.Workbook) => Promise<void>
): Promise<Blob> {
  // Fetch the template as ArrayBuffer
  const response = await fetch(templateUrl);
  const arrayBuffer = await response.arrayBuffer();

  // Load workbook from template
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);

  // Fill in the data
  await fillData(workbook);

  // Export to Blob
  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}