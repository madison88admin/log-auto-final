import '../setupJquery';

import React, { useEffect, useState } from 'react';
import * as XLSX from 'xlsx';
import 'luckysheet/dist/css/luckysheet.css';
import luckysheet from 'luckysheet';

const LuckysheetPreview: React.FC<{ excelBlob: Blob | null }> = ({ excelBlob }) => {
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (!excelBlob) return;
    setLoading(true);
    setError(null);

    if ((window as any).luckysheet) {
      try {
        (window as any).luckysheet.destroy('luckysheet-container');
      } catch (e) {}
    }

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target!.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheetDataRaw = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });
        const sheetData = Array.isArray(sheetDataRaw)
          ? sheetDataRaw.filter(row => Array.isArray(row) && row.some(cell => cell !== null && cell !== undefined && cell !== ''))
          : [];
        // Pad all rows to the maximum column count
        const maxColCount = sheetData.reduce<number>(
          (max, row) => Array.isArray(row) ? Math.max(max, row.length) : max,
          0
        );
        const paddedSheetData = sheetData.map(row =>
          Array.isArray(row)
            ? [...row, ...Array(maxColCount - row.length).fill('')]
            : Array(maxColCount).fill('')
        );
        console.log('Luckysheet sheet names:', workbook.SheetNames);
        console.log('All non-empty rows for Luckysheet (first 10 rows):', paddedSheetData.slice(0, 10));
        const rowCount = paddedSheetData.length;
        const colCount = maxColCount;

        // Defensive check for empty data
        if (!Array.isArray(paddedSheetData) || rowCount === 0 || colCount === 0) {
          setError('No data to preview. The sheet is empty or invalid.');
          setLoading(false);
          return;
        }

        luckysheet.create({
          container: 'luckysheet-container',
          data: [{
            name: 'Report',
            data: paddedSheetData,
            row: rowCount,
            column: colCount,
            order: 0,
            index: 0,
            status: 1
          }],
          showinfobar: false,
          allowEdit: false,
          showtoolbar: false,
          showstatisticBar: false,
          showsheetbar: false,
          lang: 'en',
          gridKey: Date.now().toString() + Math.random(),
          hook: {
            updated: () => setLoading(false),
            mounted: () => setLoading(false),
          },
        });

        setTimeout(() => setLoading(false), 1000);
      } catch (err) {
        setLoading(false);
        setError('Failed to preview Excel file. ' + (err instanceof Error ? err.message : ''));
        console.error('Luckysheet preview error:', err);
      }
    };
    reader.readAsArrayBuffer(excelBlob);

    return () => {
      if ((window as any).luckysheet) {
        try {
          (window as any).luckysheet.destroy('luckysheet-container');
        } catch (e) {}
      }
    };
  }, [excelBlob]);

  return (
    <div style={{ position: 'relative', width: '100%', minHeight: 600, background: '#fff', borderRadius: 8, boxShadow: '0 2px 8px rgba(0,0,0,0.08)', border: '1px solid #e5e7eb', overflow: 'hidden', margin: '0 auto', maxWidth: 1200 }}>
      {loading && (
        <div style={{
          position: 'absolute', left: 0, top: 0, width: '100%', height: '100%',
          display: 'flex', alignItems: 'center', justifyContent: 'center', background: 'rgba(255,255,255,0.7)', zIndex: 10
        }}>
          <span style={{ fontSize: 18, color: '#555' }}>Loading preview...</span>
        </div>
      )}
      {error && (
        <div style={{
          position: 'absolute', left: 0, top: 0, width: '100%', height: '100%',
          display: 'flex', alignItems: 'center', justifyContent: 'center', background: 'rgba(255,0,0,0.1)', zIndex: 11
        }}>
          <span style={{ fontSize: 18, color: '#b91c1c' }}>{error}</span>
        </div>
      )}
      <div id="luckysheet-container" style={{ width: '100%', minHeight: 600 }} />
    </div>
  );
};

export default LuckysheetPreview; 