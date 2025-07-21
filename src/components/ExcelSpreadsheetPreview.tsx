import React, { useEffect, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import Spreadsheet from 'x-data-spreadsheet';
import 'x-data-spreadsheet/dist/xspreadsheet.css';

export type ExcelInput = File | ArrayBuffer | XLSX.WorkBook | null;

interface ExcelSpreadsheetPreviewProps {
  excelFile: ExcelInput;
  title?: string;
}

function extractFirstTableWithHeader(rows: any[][], headerKeyword: string = 'CASE NOS') {
  // Find all header row indices
  const headerIndices = rows
    .map((row, idx) => (row[0] && row[0].toString().toUpperCase().includes(headerKeyword) ? idx : -1))
    .filter(idx => idx !== -1);
  if (headerIndices.length === 0) return [];
  const firstHeaderIdx = headerIndices[0];
  // Find the next header or end of data
  const nextHeaderIdx = headerIndices[1] || rows.length;
  return rows.slice(firstHeaderIdx, nextHeaderIdx);
}

function sheetjsToXSheetDataTableOnly(ws: XLSX.WorkSheet): any {
  const allRows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false });
  const tableRows = extractFirstTableWithHeader(allRows as any[][], 'CASE NOS');
  if (!tableRows.length) return { name: 'Sheet1', rows: {} };
  const rows = tableRows.length;
  const cols = Array.isArray(tableRows[0]) ? tableRows[0].length : 0;
  const data: any = { name: 'Sheet1', rows: {} };
  for (let r = 0; r < rows; ++r) {
    data.rows[r] = { cells: {} };
    const arr = tableRows[r] as any[];
    for (let c = 0; c < cols; ++c) {
      const v = arr[c];
      if (v !== undefined && v !== null && v !== '') {
        data.rows[r].cells[c] = { text: String(v) };
        if (r === 0) {
          data.rows[r].cells[c].bold = true;
        }
      }
    }
  }
  return data;
}

export const ExcelSpreadsheetPreview: React.FC<ExcelSpreadsheetPreviewProps> = ({ excelFile, title }) => {
  const containerRef = useRef<HTMLDivElement>(null);
  const spreadsheetRef = useRef<any>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let destroyed = false;
    async function renderSheet() {
      setError(null);
      if (!excelFile || !containerRef.current) return;
      let wb: XLSX.WorkBook;
      try {
        if (excelFile instanceof File) {
          const arrayBuffer = await excelFile.arrayBuffer();
          wb = XLSX.read(arrayBuffer, { type: 'array' });
        } else if (excelFile instanceof ArrayBuffer) {
          wb = XLSX.read(excelFile, { type: 'array' });
        } else if (excelFile) {
          wb = excelFile;
        } else {
          return;
        }
        // Debug: log workbook and sheet names
        console.log('Excel preview workbook:', wb);
        console.log('Sheet names:', wb.SheetNames);
        if (!wb.SheetNames.length) {
          setError('No sheets found in Excel file.');
          return;
        }
        // Always show the second sheet if it exists, otherwise the first
        const sheetName = wb.SheetNames[1] || wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        // Debug: log worksheet data
        console.log('Previewing sheet:', sheetName, ws);
        if (!ws || Object.keys(ws).length === 0) {
          setError('First sheet is empty or could not be parsed.');
          return;
        }
        // Replace sheetjsToXSheetData with the new function
        const data = sheetjsToXSheetDataTableOnly(ws);
        // Debug: log parsed data
        console.log('Parsed sheet data for preview:', data);
        if (!data || !data.rows || Object.keys(data.rows).length === 0) {
          setError('Sheet data is empty.');
          return;
        }
        if (spreadsheetRef.current && typeof spreadsheetRef.current.destroy === 'function') {
          spreadsheetRef.current.destroy();
        } else if (containerRef.current) {
          containerRef.current.innerHTML = '';
        }
        spreadsheetRef.current = new Spreadsheet(containerRef.current, { mode: 'read', showToolbar: false, showBottomBar: false })
          .loadData({ name: sheetName, freeze: 'A1', styles: [], sheets: [data] })
          .change(() => {});
      } catch (e: any) {
        setError('Failed to parse Excel file: ' + (e.message || e.toString()));
        // For debugging
        // eslint-disable-next-line no-console
        console.error('Excel preview error:', e);
      }
    }
    renderSheet();
    return () => {
      destroyed = true;
      if (spreadsheetRef.current && typeof spreadsheetRef.current.destroy === 'function') {
        spreadsheetRef.current.destroy();
      } else if (containerRef.current) {
        containerRef.current.innerHTML = '';
      }
      spreadsheetRef.current = null;
    };
  }, [excelFile]);

  return (
    <div style={{ height: 520, width: '100%', background: '#fff', borderRadius: 8, border: '1px solid #eee', overflow: 'hidden' }}>
      {title && <div style={{ fontWeight: 600, padding: '8px 16px', borderBottom: '1px solid #eee', background: '#fafbfc' }}>{title}</div>}
      {error ? (
        <div style={{ color: 'red', padding: 16 }}>{error}</div>
      ) : (
        <div ref={containerRef} style={{ height: 480, width: '100%', overflow: 'auto' }} />
      )}
    </div>
  );
}; 