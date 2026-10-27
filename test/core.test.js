// 核心邏輯測試：node --test test/
const test = require('node:test');
const assert = require('node:assert');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const JSZip = require('jszip');
const Core = require('../js/core.js');

// 載入 template/template.js（瀏覽器用，掛在 window 上）
global.window = {};
eval(fs.readFileSync(path.join(__dirname, '..', 'template', 'template.js'), 'utf8'));
const TPL = global.window.DOCX_TEMPLATE;

function setup(cfgPatch) {
  const wb = XLSX.readFile(path.join(__dirname, 'fixtures', 'sample.xlsx'), { cellNF: true });
  const t = Core.sheetToTable(XLSX, wb.Sheets[wb.SheetNames[0]]);
  const map = Core.autoMap(t.headers, 115);
  const recs = Core.buildRecords(t, map, {});
  const cfg = Object.assign({ total: 150, customs: [], excluded: [], autoClassify: true, th: { y15: 15, y5: 5, Y1: 114, Y2: 113 } }, cfgPatch || {});
  Core.assign(recs, cfg); Core.syncCats(recs, cfg);
  return { t, map, recs, cfg, sel: Core.select(recs, cfg) };
}

test('自動對應範例欄位', () => {
  const { map } = setup();
  assert.strictEqual(map.id, '溪流編號');
  assert.strictEqual(map.town, '鄉鎮');
  assert.strictEqual(map.ev1, '114年事件崩塌率');
  assert.strictEqual(map.ev2, '113年事件崩塌率');
  assert.strictEqual(map.wsr, '集水區崩塌率(%)');
  assert.strictEqual(map.order, '發文排序');
});

test('百分比格式與 Word 一致', () => {
  assert.strictEqual(Core.fmtPct(0.15), '0.15%');
  assert.strictEqual(Core.fmtPct(0.005), '0.005%');
  assert.strictEqual(Core.fmtPct(0.002), '0.002%');
  assert.strictEqual(Core.fmtPct(0), '0.00%');
  assert.strictEqual(Core.fmtPct(null), '－');
});

test('篩選類別去除排序前綴', () => {
  assert.strictEqual(Core.cleanCat('0地方政府提報'), '地方政府提報');
  assert.strictEqual(Core.cleanCat('第1類'), '第1類');
});

test('預設優先順序與空白類別自動判定', () => {
  const { cfg, recs } = setup();
  assert.deepStrictEqual(cfg.cats.map(c => c.label), ['地方政府提報', '農村水保署主動提報', '第1類', '第2類', '第3類', '第4類']);
  const r40 = recs.find(r => r.id === '範市DF040'); assert.strictEqual(r40.catLabel, '第1類');
  const r41 = recs.find(r => r.id === '範市DF041'); assert.strictEqual(r41.catLabel, '第3類');
  const r42 = recs.find(r => r.id === '範市DF042'); assert.strictEqual(r42.catLabel, '第4類');
});

test('調查數量與類別上限', () => {
  const s = setup({ total: 6 });
  assert.strictEqual(s.sel.list.length, 6);
  const s2 = setup();
  s2.cfg.cats.find(c => c.label === '第2類').quota = '1';
  const sel = Core.select(s2.recs, s2.cfg);
  assert.strictEqual(sel.list.filter(r => r.catLabel === '第2類').length, 1);
});

test('業主指定類別置頂並依輸入順序', () => {
  const { sel } = setup({ customs: [{ uid: 'x', name: '陳情案', reason: '民眾陳情', ids: '範市DF031\n範縣DF010' }] });
  assert.deepStrictEqual(sel.list.slice(0, 2).map(r => r.id), ['範市DF031', '範縣DF010']);
  assert.strictEqual(sel.list[0].catLabel, '陳情案');
  assert.strictEqual(sel.list[0].reasonOut, '民眾陳情');
});

test('移出名單後自動遞補', () => {
  const { sel } = setup({ total: 3, excluded: ['範縣G114-1'] });
  assert.strictEqual(sel.list.length, 3);
  assert.ok(!sel.list.some(r => r.id === '範縣G114-1'));
});

test('註記文字', () => {
  const { sel, cfg } = setup();
  const n = Core.noteText(sel, cfg);
  assert.match(n, /^註：項次第1~9項\(篩選類別：地方政府提報、農村水保署主動提報及第1~2類\)建議優先納入/);
  assert.match(n, /則以項次第10~14項\(第3~4類\)依序列入現勘名單。$/);
});

test('匯出 Word：標題、資料列、字型', async () => {
  const { sel, cfg } = setup();
  const buf = await Core.buildDocx(JSZip, TPL.b64, TPL.row, sel.list, { title: '測試清冊', Y: 115, note: Core.noteText(sel, cfg), nodeBuffer: true });
  const zip = await JSZip.loadAsync(buf);
  const doc = await zip.file('word/document.xml').async('string');
  assert.ok(doc.includes('測試清冊'));
  assert.ok(!/\{\{[A-Z0-9]+\}\}/.test(doc), '不得殘留佔位字');
  assert.strictEqual((doc.match(/<w:tr[ >]/g) || []).length, 2 + sel.list.length);
  assert.ok(doc.includes('w:eastAsia="標楷體"') && doc.includes('w:ascii="Times New Roman"'));
  assert.ok(!/w:rFonts(?![^>]*標楷體)[^>]*\/>/.test(doc), '所有 rFonts 皆為標楷體／Times New Roman');
  assert.ok(doc.includes('<w:t>114</w:t>') && doc.includes('<w:t>113</w:t>'));
});
