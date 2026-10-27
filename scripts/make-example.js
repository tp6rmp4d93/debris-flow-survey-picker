// 產生虛構的範例試算表（examples/範例_候選溪流.xlsx 與 test/fixtures/sample.xlsx）
// 用法：node scripts/make-example.js
const XLSX = require('xlsx');
const path = require('path');
const H = ['項次', '發文排序', '溪流編號', '縣市', '鄉鎮', '村里', '篩選類別', '初步篩選辦理原因', '事件名稱',
  '114年事件崩塌率', '113年事件崩塌率', '集水區崩塌率(%)', '風險等級', '114年紅色警戒次數', '距前次調查年', '保全住戶數量'];
const D = '－';
// 全部為虛構資料，僅供操作示範
const rows = [
  [1, 1, '範縣G114-1', '範例縣', '甲鄉', '一村', '0地方政府提報', '新增疑似土石流災害', null, D, D, D, D, D, D, D],
  [2, 2, '範縣DF001', '範例縣', '甲鄉', '二村', '0地方政府提報', '申請調整風險等級', null, D, D, D, '中', 1, 12, 8],
  [3, 3, '範縣DF002', '範例縣', '乙鎮', '三里', '0農村水保署主動提報', '申請調整風險等級', null, 0.0008, D, 0.0012, '持續觀察', 2, 7, D],
  [4, 4, '範縣DF010', '範例縣', '乙鎮', '四里', '第1類', '距前次更新已逾15年且具保全住戶', null, D, D, 0.0709, '低', D, 16, 3],
  [5, 5, '範縣DF011', '範例縣', '丙區', '五里', '第1類', '距前次更新已逾15年且具保全住戶', null, D, D, D, '高', D, 17, 15],
  [6, 6, '範市DF020', '範例市', '丁區', '六里', '第2類', '114年度颱風豪雨事件致災', '範例颱風', D, 0.0008, 0.0286, '高', 1, 1, 17],
  [7, 7, '範市DF021', '範例市', '丁區', '七里', '第2類', '114年度集水區具新生崩塌地', null, 0.00002, 0.0004, 0.0638, '高', 2, 1, 26],
  [8, 8, '範市DF022', '範例市', '戊區', '八里', '第2類', '113年度集水區具新生崩塌地', null, D, 0.0005, 0.0073, '低', 1, 15, 0],
  [9, 9, '範市DF030', '範例市', '戊區', '九里', '第3類', '114年度發布紅色警戒且具前次更新滿5年', null, D, D, D, '高', 1, 13, 78],
  [10, 10, '範市DF031', '範例市', '己鄉', '十村', '第3類', '114年度發布紅色警戒且具前次更新滿5年', null, D, D, 0.012, '中', 2, 9, 4],
  [11, 11, '範市DF032', '範例市', '己鄉', '十一村', '第3類', '114年度發布紅色警戒且具前次更新滿5年', null, D, D, D, '低', 1, 6, 2],
  [12, null, '範市DF040', '範例市', '庚鄉', '十二村', null, null, null, D, D, D, '中', D, 20, 5],
  [13, null, '範市DF041', '範例市', '庚鄉', '十三村', null, null, null, D, D, D, '高', 3, 8, 12],
  [14, null, '範市DF042', '範例市', '辛鄉', '十四村', null, null, null, D, D, D, '低', D, 4, 1]
];
const ws = XLSX.utils.aoa_to_sheet([H].concat(rows));
// 崩塌率欄位設百分比格式（與實際資料相同）
const range = XLSX.utils.decode_range(ws['!ref']);
for (let r = 1; r <= range.e.r; r++) for (const c of [9, 10, 11]) {
  const cell = ws[XLSX.utils.encode_cell({ r, c })]; if (cell && cell.t === 'n') cell.z = '0.00%';
}
const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, '候選溪流');
const root = path.join(__dirname, '..');
XLSX.writeFile(wb, path.join(root, 'examples', '範例_候選溪流.xlsx'));
XLSX.writeFile(wb, path.join(root, 'test', 'fixtures', 'sample.xlsx'));
console.log('已產生範例試算表');
