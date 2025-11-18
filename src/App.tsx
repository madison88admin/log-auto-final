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
  const [parsedTables, setParsedTables] = useState<{ table: any[][], modelName: string, mergedCellText?: string }[]>([]);  const [selectedTableIdx, setSelectedTableIdx] = useState(0);
  // Add state to store all generated report buffers
  const [reportBuffers, setReportBuffers] = useState<ArrayBuffer[]>([]);
  // Add state to store sheet names for download
  const [reportSheetNames, setReportSheetNames] = useState<string[]>([]);
  // Add state for merged cell preview
  const [mergedCellText, setMergedCellText] = useState<string>('');
  // 1. Add state variable for backend-generated Blob
  // const [generatedReportBlob, setGeneratedReportBlob] = useState<Blob | null>(null);

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
      let poNo = idxPoNo !== -1 ? String(dataRows[0]?.[idxPoNo] || `Report ${t+1}`) : `Report ${t+1}`;
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
            modelName = safe(table[i-2], idxCaseNos);
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
        const cartonRowRanges: Record<string, {start: number, end: number}> = {};
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
          // A=0, B=1, ..., L=11, M=12, N=13, O=14, P=15, Q=16, R=17, S=18, T=19, U=20
          // ws.getCell(`G${rowNum}`).value = safe(row, 11); // G = L (replaced by split-carton logic above)
          ws.getCell(`N${rowNum}`).value = safe(row, 10); // N = K
          ws.getCell(`O${rowNum}`).value = safe(row, 12); // O = M
          ws.getCell(`P${rowNum}`).value = safe(row, 14); // P = O
          ws.getCell(`Q${rowNum}`).value = safe(row, 16); // Q = Q
          ws.getCell(`R${rowNum}`).value = safe(row, 17); // R = R
          ws.getCell(`S${rowNum}`).value = safe(row, 18); // S = S
          ws.getCell(`T${rowNum}`).value = safe(row, 19); // T = T
          ws.getCell(`U${rowNum}`).value = safe(row, 20); // U = U
          ws.getCell(`V${rowNum}`).value = safe(row, 20); // V = U
        });

        // 5. Merge Carton# cells for each group in the worksheet
        Object.values(cartonRowRanges).forEach(({start, end}) => {
          if (end > start) {
            ws.mergeCells(`C${start}:C${end}`);
            // Also merge columns N–V for this group (Column N is 14, O=15, ..., V=22)
            const colLetters = ['N','O','P','Q','R','S','T','U','V'];
            colLetters.forEach(col => {
              ws.mergeCells(`${col}${start}:${col}${end}`);
            });
          }
        });

        // --- ROUNDING LOGIC FOR P, Q, U, V (16, 17, 21, 22) ---
        const roundingCols = [16, 17, 21, 22];
        for (let i = 0; i < numDataRows; i++) {
          const rowNum = mainTableStart + i;
          roundingCols.forEach(colIdx => {
            const cell = ws.getRow(rowNum).getCell(colIdx);
            if (typeof cell.value === 'number' && !isNaN(cell.value)) {
              cell.value = Math.round(cell.value * 1000) / 1000;
            }
          });
        }

        // --- TOTAL CARTON LOGIC: sum only the first non-empty cell of each merged group in Column N ---
        let totalCarton = 0;
        let prevCartonVal = null;
        let totalNetWeight = 0;
        let totalGrossWeight = 0;
        let totalCBM = 0;
        for (let i = 0; i < numDataRows; i++) {
          const rowNum = mainTableStart + i;
          const cartonVal = ws.getRow(rowNum).getCell(14).value;
          if (
            cartonVal &&
            cartonVal !== '' &&
            cartonVal !== 'nan' &&
            cartonVal !== 'none' &&
            cartonVal !== prevCartonVal
          ) {
            totalCarton += Number(cartonVal);
            // Only sum these columns for the first row of each merged group
            totalNetWeight += parseFloat(ws.getCell(`P${rowNum}`).value as string) || 0;
            totalGrossWeight += parseFloat(ws.getCell(`Q${rowNum}`).value as string) || 0;
            totalCBM += parseFloat(ws.getCell(`V${rowNum}`).value as string) || 0;
            prevCartonVal = cartonVal;
          }
        }
        totalCarton = Math.floor(totalCarton);

        // --- SUMMARY AND COLOR BREAKDOWN ---
        // Move summary and color breakdown to start 1 row below the last data row
        const summaryStartRow = mainTableStart + numDataRows + 1;
        ws.mergeCells(`D${summaryStartRow}:E${summaryStartRow}`);
        ws.getCell(`D${summaryStartRow}`).value = 'Summary';
        ws.getCell(`D${summaryStartRow}`).alignment = { horizontal: 'center', vertical: 'middle' };

        // Write summary titles in D, values in E
        ws.getCell(`D${summaryStartRow + 1}`).value = 'Total Carton';
        ws.getCell(`E${summaryStartRow + 1}`).value = totalCarton;
        ws.getCell(`D${summaryStartRow + 2}`).value = 'Total Net Weight';
        ws.getCell(`E${summaryStartRow + 2}`).value = Math.round(totalNetWeight * 1000) / 1000;
        ws.getCell(`D${summaryStartRow + 3}`).value = 'Total Gross Weight';
        ws.getCell(`E${summaryStartRow + 3}`).value = Math.round(totalGrossWeight * 1000) / 1000;
        ws.getCell(`D${summaryStartRow + 4}`).value = 'Total CBM';
        ws.getCell(`E${summaryStartRow + 4}`).value = Math.round(totalCBM * 1000) / 1000;

        // --- COLOR BREAKDOWN: multiply size values by propagated Units/CRT from Column N ---
        const sizeColIndices = [7, 8, 9, 10, 11, 12, 13]; // G–M
        const colorColIdx = 4; // D
        const unitsCrtColIdx = 14; // N
        let colorMap: Record<string, number[]> = {};
        let lastUnitsCrt: number | null = null;
        for (let i = 0; i < numDataRows; i++) {
          const rowNum = mainTableStart + i;
          const color = ws.getRow(rowNum).getCell(colorColIdx).value?.toString().trim();
          let unitsCrt = ws.getRow(rowNum).getCell(unitsCrtColIdx).value;
          if (unitsCrt && unitsCrt !== '' && unitsCrt !== 'nan' && unitsCrt !== 'none') {
            lastUnitsCrt = Number(unitsCrt);
          }
          if (!color) continue;
          if (!colorMap[color]) colorMap[color] = Array(sizeNames.length).fill(0);
          sizeColIndices.forEach((colIdx, j) => {
            const sizeVal = Number(ws.getRow(rowNum).getCell(colIdx).value) || 0;
            colorMap[color][j] += sizeVal * (lastUnitsCrt ?? 0);
          });
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

        // --- ANCHOR IMAGE TO CELL T22 ---
        // Find and reposition the existing image to anchor it to cell T22
        try {
          // Try to get images using different methods
          console.log('Worksheet object:', ws);
          console.log('Worksheet properties:', Object.keys(ws));
          
          // Check if getImages method exists
          if (typeof ws.getImages === 'function') {
            const images = ws.getImages();
            console.log('Found images using getImages():', images ? images.length : 0);
            
            if (images && images.length > 0) {
              console.log('First image:', images[0]);
              console.log('Image properties:', Object.keys(images[0]));
            }
          } else {
            console.log('getImages() method not available');
          }
          
          // Alternative: check if images are stored differently
          // if (ws.images) {
          //   console.log('ws.images found:', ws.images);
          // }
          
          // Alternative: check if images are in the model
          // if (ws.model && ws.model.images) {
          //   console.log('ws.model.images found:', ws.model.images);
          // }
          
        } catch (error) {
          console.error('Error handling image anchoring:', error);
          if (error instanceof Error) {
            console.error('Error stack:', error.stack);
          }
          // Continue without image anchoring if there's an error
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

  // 2. In handleExportAllAsSingleExcel, set the Blob
  const handleExportAllAsSingleExcel = async () => {
    if (!uploadedFile) return;
    setIsProcessing(true);
    setError(null);
    try {
      const formData = new FormData();
      formData.append('file', uploadedFile);

      const API_URL = import.meta.env.VITE_API_URL || 'https://log-auto-final-python.onrender.com';
      console.log('API_URL:', API_URL);
      const response = await fetch(`${API_URL}/generate-reports/`, {
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

  // 4. Reset the Blob in handleReset
  const handleReset = () => {
    setUploadedFile(null);
    setProcessingResult(null);
    setGeneratedExcelBuffer(null);
    setError(null);
    setIsProcessing(false);
    setParsedTables([]);
    setSelectedTableIdx(0);
    setMergedCellText('');
    // setGeneratedReportBlob(null); // <-- Reset backend-generated Blob
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
                        <svg width="20" height="20" fill="none" stroke="#2563eb" strokeWidth="2" viewBox="0 0 24 24" style={{ marginRight: 6 }}><rect x="4" y="4" width="16" height="16" rx="4"/><path d="M8 9h8M8 13h5"/></svg>
                        <b style={{ fontWeight: 600, color: '#2563eb', fontSize: 15 }}>Ship to:</b>
                      </span>
                      <span style={{ flex: 1, fontWeight: 500 }}>{mergedCellText}</span>
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