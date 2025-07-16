import React, { useMemo, useState } from 'react';
import * as XLSX from 'xlsx';
import { DataGrid, GridColDef } from '@mui/x-data-grid';
import { Box, Select, MenuItem, Typography } from '@mui/material';

export type ExcelInput = File | ArrayBuffer | XLSX.WorkBook;

interface ExcelPreviewGridProps {
  excelFile: ExcelInput | null;
  title?: string;
}

export const ExcelPreviewGrid: React.FC<ExcelPreviewGridProps> = ({ excelFile, title }) => {
  const [sheetName, setSheetName] = useState<string | null>(null);
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);

  React.useEffect(() => {
    if (!excelFile) {
      setWorkbook(null);
      setSheetName(null);
      return;
    }
    const load = async () => {
      let wb: XLSX.WorkBook;
      if (excelFile instanceof File) {
        const arrayBuffer = await excelFile.arrayBuffer();
        wb = XLSX.read(arrayBuffer, { type: 'array' });
      } else if (excelFile instanceof ArrayBuffer) {
        wb = XLSX.read(excelFile, { type: 'array' });
      } else {
        wb = excelFile;
      }
      setWorkbook(wb);
      setSheetName(wb.SheetNames[0] || null);
    };
    load();
  }, [excelFile]);

  const sheetData = useMemo(() => {
    if (!workbook || !sheetName) return { columns: [], rows: [] };
    const ws = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(ws, { header: 1 });
    if (!json.length) return { columns: [], rows: [] };
    const headers = (json[0] as string[]).map((h, i) => h || `Column ${i + 1}`);
    const columns: GridColDef[] = headers.map((header, i) => ({
      field: `col_${i}`,
      headerName: header,
      width: 150,
      flex: 1,
      sortable: false,
      filterable: false,
    }));
    const rows = json.slice(1, 101).map((row, idx: number) => {
      const arr = row as any[];
      const rowObj: any = { id: idx };
      headers.forEach((_, i) => {
        rowObj[`col_${i}`] = arr[i] ?? '';
      });
      return rowObj;
    });
    return { columns, rows };
  }, [workbook, sheetName]);

  if (!excelFile) {
    return (
      <Box p={2} textAlign="center" color="text.secondary">
        <Typography variant="body2">No file selected</Typography>
      </Box>
    );
  }

  return (
    <Box p={2} height="100%" display="flex" flexDirection="column">
      <Box mb={1} display="flex" alignItems="center" justifyContent="space-between">
        <Typography variant="subtitle1" fontWeight={600}>{title}</Typography>
        {workbook && (
          <Select
            size="small"
            value={sheetName || ''}
            onChange={e => setSheetName(e.target.value)}
            sx={{ minWidth: 120 }}
          >
            {workbook.SheetNames.map(name => (
              <MenuItem key={name} value={name}>{name}</MenuItem>
            ))}
          </Select>
        )}
      </Box>
      <Box flex={1} minHeight={300}>
        <DataGrid
          columns={sheetData.columns}
          rows={sheetData.rows}
          autoHeight={false}
          density="compact"
          hideFooterSelectedRowCount
          hideFooterPagination
          sx={{ height: '100%', border: 1, borderColor: 'divider', background: '#fff' }}
        />
      </Box>
    </Box>
  );
}; 