const XLSX = require('xlsx');
const wb = XLSX.readFile('data.xlsx');
const ws = wb.Sheets['数据统计'];
console.log('!ref:', ws['!ref']);
// Print all non-empty cells in rows 94-106
for (const key of Object.keys(ws)) {
  if (key.startsWith('!')) continue;
  const cell = ws[key];
  const m = /^([A-Z]+)(\d+)$/.exec(key);
  if (!m) continue;
  const row = parseInt(m[2], 10);
  if (row >= 94 && row <= 108) {
    console.log(`${key}=t:${cell.t} v:${JSON.stringify(cell.v).slice(0,50)} f:${cell.f||''}`);
  }
}
