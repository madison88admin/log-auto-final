import React, { useEffect, useState } from 'react';
import * as XLSX from 'xlsx';
import { HotTable } from '@handsontable/react';
import 'handsontable/dist/handsontable.full.min.css';

interface ExcelHandsontablePreviewProps {
  excelFile?: File | Blob | ArrayBuffer | null;
  title?: string;
  onTablesExtracted?: (tables: { table: any[][], modelName: string, mergedCellText?: string }[]) => void;
  onSelectedTableChange?: (table: any[][], modelName: string) => void;
}

// Levenshtein distance for fuzzy header matching
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

function normalizeHeader(header: string) {
  return header.replace(/\s+/g, ' ').trim().toLowerCase();
}

function fuzzyHeaderMatch(row: string[], expected: string[]): boolean {
  if (row.length < expected.length) return false;
  let matches = 0;
  console.log('Fuzzy matching row:', row);
  console.log('Against expected:', expected);
  
  for (let i = 0; i < expected.length; i++) {
    const cell = normalizeHeader(row[i] || '');
    const exp = normalizeHeader(expected[i]);
    const distance = levenshtein(cell, exp);
    const isMatch = cell === exp || distance <= 2;
    
    console.log(`  Column ${i}: "${cell}" vs "${exp}" (distance: ${distance}, match: ${isMatch})`);
    
    if (isMatch) {
      matches++;
    }
  }
  
  const matchRatio = matches / expected.length;
  console.log(`Match ratio: ${matches}/${expected.length} = ${matchRatio}`);
  
  // At least 90% of columns must match
  return matchRatio >= 0.9;
}

const EXPECTED_HEADER = [
  'CASE NOS', 'SA4 PO NO#', 'CARTON PO NO#', 'SAP STYLE NO', 'STYLE NAME #',
  'S4 Material', 'Material No#', 'COLOR', 'Size', 'Total QTY', 'CARTON',
  'QTY / CARTON', 'TOTAL QTY', 'N.N.W / ctn', 'TOTAL N.N.W.', 'N.W / ctn', 'TOTAL N.W.', 'G.W. / ctn',
  'TOTAL G.W.', 'MEAS. CM', 'TOTAL CBM'
];

export const ExcelHandsontablePreview: React.FC<ExcelHandsontablePreviewProps> = ({ excelFile, title, onTablesExtracted, onSelectedTableChange }) => {
  const [tables, setTables] = useState<{ table: any[][], modelName: string, mergedCellText?: string }[]>([]);
  const [selectedTableIdx, setSelectedTableIdx] = useState(0);

  useEffect(() => {
    if (!excelFile) {
      setTables([]);
      setSelectedTableIdx(0);
      if (onTablesExtracted) onTablesExtracted([]);
      return;
    }
    (async () => {
      try {
        console.log('Processing excelFile:', excelFile);
        let arrayBuffer: ArrayBuffer;
        if (excelFile instanceof File || excelFile instanceof Blob) {
          arrayBuffer = await excelFile.arrayBuffer();
        } else if (excelFile instanceof ArrayBuffer) {
          arrayBuffer = excelFile;
        } else {
          console.log('Unsupported file type:', typeof excelFile);
          setTables([]);
          setSelectedTableIdx(0);
          if (onTablesExtracted) onTablesExtracted([]);
          return;
        }
        
        const wb = XLSX.read(arrayBuffer, { type: 'array' });
        console.log('Workbook sheets:', wb.SheetNames);
        
        // Always use the second sheet if available
        const sheetName = wb.SheetNames[1] || wb.SheetNames[0];
        console.log('Using sheet:', sheetName);
        
        const ws = wb.Sheets[sheetName];
        const allRows = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];
        console.log('Total rows in sheet:', allRows.length);
        console.log('First few rows:', allRows.slice(0, 5));

        // --- Extract C15-C20 (row 14-19, col 2) for F6-G16 merged cell preview ---
        let mergedCellText = '';
        try {
          const mergedLines: string[] = [];
          for (let rowIdx = 14; rowIdx <= 19; rowIdx++) {
            if (allRows[rowIdx] && allRows[rowIdx][2] !== undefined && allRows[rowIdx][2] !== null && String(allRows[rowIdx][2]).trim() !== '') {
              mergedLines.push(String(allRows[rowIdx][2]).trim());
            }
          }
          mergedCellText = mergedLines.join('\n');
        } catch (e) {
          console.warn('Could not extract C15-C20 for merged cell preview:', e);
        }
        
        // Find all tables with fuzzy header match
        const foundTables: { table: any[][], modelName: string, mergedCellText?: string }[] = [];
        for (let i = 0; i < allRows.length; i++) {
          const row = (allRows[i] || []).map((cell: any) => (cell || '').toString());
          console.log(`Row ${i}:`, row);
          
          if (fuzzyHeaderMatch(row, EXPECTED_HEADER)) {
            console.log(`Found matching header at row ${i}:`, row);
            // Extract Model Name 2 rows above header (first column)
            let modelName = '';
            if (i >= 2) {
              modelName = (allRows[i - 2][0] || '').toString();
            }
            // Found a header, extract table
            const table: any[][] = [row];
            let j = i + 1;
            while (
              j < allRows.length &&
              allRows[j].some(cell => cell !== null && cell !== undefined && cell !== '') &&
              !fuzzyHeaderMatch((allRows[j] || []).map((cell: any) => (cell || '').toString()), EXPECTED_HEADER)
            ) {
              table.push(allRows[j]);
              j++;
            }
            // Pad all rows to the header's column count
            const headerLen = row.length;
            const paddedTable = table.map((r: any[]) => [...r, ...Array(headerLen - r.length).fill('')]);
            foundTables.push({ table: paddedTable, modelName, mergedCellText });
            console.log(`Extracted table with ${paddedTable.length} rows, model: ${modelName}`);
            i = j - 1; // Skip to end of this table
          }
        }
        
        console.log('Found tables:', foundTables.length);
        setTables(foundTables);
        setSelectedTableIdx(0);
        if (onTablesExtracted) onTablesExtracted(foundTables);
      } catch (e) {
        console.error('Error processing Excel file:', e);
        setTables([]);
        setSelectedTableIdx(0);
        if (onTablesExtracted) onTablesExtracted([]);
      }
    })();
  }, [excelFile]);

  useEffect(() => {
    if (onSelectedTableChange && tables[selectedTableIdx]) {
      onSelectedTableChange(tables[selectedTableIdx].table, tables[selectedTableIdx].modelName);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedTableIdx, tables]);

  if (!excelFile) return null;

  const rawTableData = tables[selectedTableIdx]?.table || [];
  const modelName = tables[selectedTableIdx]?.modelName || '';
  const mergedCellText = tables[selectedTableIdx]?.mergedCellText || '';
  
  // Clean the table data to ensure it's compatible with Handsontable
  const tableData = rawTableData.map(row => 
    row.map(cell => {
      // Convert undefined/null to empty string
      if (cell === undefined || cell === null) return '';
      // Convert to string and trim
      return String(cell).trim();
    })
  );
  
  console.log('Rendering table:', {
    selectedTableIdx,
    totalTables: tables.length,
    rawTableDataLength: rawTableData.length,
    cleanedTableDataLength: tableData.length,
    firstRow: tableData[0],
    modelName
  });

  // Calculate column widths based on header name length
  const headerRow = tableData[0] || [];
  const colWidths = headerRow.map((header: string) => {
    const base = 10; // px per character
    const min = 60;
    const max = 300;
    const width = Math.max(min, Math.min(max, (header ? header.length * base : min)));
    return width;
  });

  return (
    <div style={{ width: '100%', background: '#fff', borderRadius: 8, border: '1px solid #eee', overflow: 'hidden' }}>
      {title && <div style={{ fontWeight: 600, padding: '8px 16px', borderBottom: '1px solid #eee', background: '#fafbfc' }}>{title}</div>}
      {/* Remove F6-G16 merged cell preview from the left preview */}
      {tables.length > 1 && (
        <div style={{ padding: '8px 16px', background: '#f5f7fa', borderBottom: '1px solid #eee', display: 'flex', alignItems: 'center', gap: 8 }}>
          <span style={{ fontWeight: 500 }}>Select Table:</span>
          <select value={selectedTableIdx} onChange={e => setSelectedTableIdx(Number(e.target.value))}>
            {tables.map((t, idx) => (
              <option key={idx} value={idx}>Table {idx + 1} {t.modelName ? `- Model: ${t.modelName}` : ''}</option>
            ))}
          </select>
          <span style={{ color: '#888', fontSize: 12 }}>({tables.length} detected)</span>
        </div>
      )}
      {modelName && (
        <div style={{ padding: '8px 16px', background: '#f9fafb', borderBottom: '1px solid #eee', fontStyle: 'italic', color: '#444' }}>
          Model Name: {modelName}
        </div>
      )}
      {tableData.length > 1 ? (
        <>
          {console.log('Handsontable data:', tableData.slice(1))}
          {console.log('Handsontable headers:', tableData[0])}
          <HotTable
            data={tableData.slice(1)} // skip header row
            colHeaders={tableData[0]} // use first row as headers
            rowHeaders={true}
            width="100%"
            height="400px"
            licenseKey="non-commercial-and-evaluation"
            stretchH="all"
            colWidths={colWidths}
            afterInit={() => console.log('Handsontable initialized with data:', tableData.length, 'rows')}
            afterRender={() => console.log('Handsontable rendered')}
          />
        </>
      ) : (
        <div style={{ padding: '20px', textAlign: 'center', color: '#666' }}>
          <p>No table data available</p>
          <p>Tables found: {tables.length}</p>
          <p>Selected table index: {selectedTableIdx}</p>
          <p>Raw table data length: {rawTableData.length}</p>
          <p>Cleaned table data length: {tableData.length}</p>
        </div>
      )}
    </div>
  );
}; 