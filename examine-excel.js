const XLSX = require('xlsx');

console.log('=== Examining ARC_ReportGenerator COPY.xlsm ===');
try {
  const arcWorkbook = XLSX.readFile('public/ARC_ReportGenerator COPY.xlsm');
  console.log('Sheets:', arcWorkbook.SheetNames);
  
  const arcReportSheet = arcWorkbook.Sheets['Report'];
  const arcData = XLSX.utils.sheet_to_json(arcReportSheet, {header: 1});
  console.log('ARC Report sheet - First 20 rows:');
  arcData.slice(0, 20).forEach((row, i) => {
    console.log(`Row ${i+1}:`, row);
  });
  
  console.log('\n=== Examining ReportTemplate.xlsx ===');
  const templateWorkbook = XLSX.readFile('public/ReportTemplate.xlsx');
  console.log('Sheets:', templateWorkbook.SheetNames);
  
  const templateReportSheet = templateWorkbook.Sheets['Report'];
  const templateData = XLSX.utils.sheet_to_json(templateReportSheet, {header: 1});
  console.log('Template Report sheet - First 20 rows:');
  templateData.slice(0, 20).forEach((row, i) => {
    console.log(`Row ${i+1}:`, row);
  });
  
} catch(e) {
  console.log('Error:', e.message);
} 