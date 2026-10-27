/* 土石流潛勢溪流現勘名單篩選 — 介面程式
 * 依賴：vendor/xlsx.full.min.js、vendor/jszip.min.js、template/template.js、js/core.js、js/version.js
 */
(function () {
'use strict';
var VERSION = window.APP_VERSION || 'dev';
var TPL = window.DOCX_TEMPLATE;
var LS = 'df-survey-picker-v1';
var COLORS = ['--s1', '--s2', '--s3', '--s4', '--s5', '--s6', '--s7', '--s8'];
var DEF_CFG = { total: 150, year: 115, customs: [], excluded: [], autoClassify: true, y15: 15, y5: 5, cats: [], pctMode: 'auto' };
var S = { table: null, fileName: '', map: {}, cfg: clone(DEF_CFG), overrides: {}, title: '', titleCustom: false, note: '', noteCustom: false, view: 'import', log: [] };
var wb = null, recs = [], sel = { list: [], perCat: {} };

function clone(o) { return JSON.parse(JSON.stringify(o)); }
function $(id) { return document.getElementById(id); }
function el(tag, attrs, kids) {
  var e = document.createElement(tag);
  if (attrs) for (var k in attrs) {
    if (k === 'text') e.textContent = attrs[k];
    else if (k.slice(0, 2) === 'on') e.addEventListener(k.slice(2), attrs[k]);
    else e.setAttribute(k, attrs[k]);
  }
  (kids || []).forEach(function (c) { if (c != null) e.appendChild(typeof c === 'string' ? document.createTextNode(c) : c); });
  return e;
}
function toast(t) { var x = $('toast'); x.textContent = t; x.classList.add('on'); clearTimeout(toast._t); toast._t = setTimeout(function () { x.classList.remove('on'); }, 2600); }
function save() { try { localStorage.setItem(LS, JSON.stringify(S)); } catch (e) { toast('瀏覽器儲存空間不足，設定未能保存'); } }
function load() {
  try {
    var s = JSON.parse(localStorage.getItem(LS) || 'null');
    if (s && s.cfg) { S = Object.assign(S, s); S.cfg = Object.assign(clone(DEF_CFG), s.cfg); S.log = s.log || []; }
  } catch (e) { /* 忽略損毀的暫存 */ }
}
function colorOf(key) { var i = S.cfg.cats.findIndex(function (c) { return c.key === key; }); return 'var(' + COLORS[(i < 0 ? 0 : i) % COLORS.length] + ')'; }
function Y() { return parseInt(S.cfg.year, 10) || 115; }
function catLabel(key) { var c = S.cfg.cats.filter(function (x) { return x.key === key; })[0]; return c ? c.label : key; }
function sortText(arr) { return (arr || []).filter(Boolean).map(function (k) { return Core.SORT_KEYS[k] ? Core.SORT_KEYS[k].label : k; }).join(' > ') || '（試算表列序）'; }

/* ---------- 變更紀錄 ----------
 * 每筆：t 時間、v 程式版本、loc 更動位置、item 更動項目、from 原值、to 新值 */
function logChange(loc, item, from, to) {
  from = from == null ? '' : String(from); to = to == null ? '' : String(to);
  if (from === to) return;
  S.log.push({ t: new Date().toISOString(), v: VERSION, loc: loc, item: item, from: from, to: to });
  if (S.log.length > 5000) S.log = S.log.slice(-5000);
}

/* ---------- 分頁 ---------- */
function show(v) {
  S.view = v;
  document.querySelectorAll('nav.steps button').forEach(function (b) { b.setAttribute('aria-selected', b.dataset.v === v ? 'true' : 'false'); });
  document.querySelectorAll('section.view').forEach(function (s) { s.classList.toggle('on', s.id === 'v-' + v); });
  if (location.hash.slice(1) !== v) history.replaceState(null, '', '#' + v);
  render(); save(); window.scrollTo(0, 0);
}
document.querySelectorAll('nav.steps button').forEach(function (b) { b.addEventListener('click', function () { show(b.dataset.v); }); });

/* ---------- 匯入 ---------- */
var drop = $('drop'), file = $('file');
drop.addEventListener('click', function () { file.click(); });
drop.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); file.click(); } });
['dragenter', 'dragover'].forEach(function (t) { drop.addEventListener(t, function (e) { e.preventDefault(); drop.classList.add('hover'); }); });
['dragleave', 'drop'].forEach(function (t) { drop.addEventListener(t, function (e) { e.preventDefault(); drop.classList.remove('hover'); }); });
drop.addEventListener('drop', function (e) { var f = e.dataTransfer.files[0]; if (f) readFile(f); });
file.addEventListener('change', function () { var f = file.files[0]; if (f) readFile(f); file.value = ''; });

function readFile(f) {
  var r = new FileReader();
  r.onload = function () {
    try { openWorkbook(new Uint8Array(r.result), f.name); }
    catch (err) { msg('impMsg', 'err', '無法讀取「' + f.name + '」：' + err.message + '。請確認檔案為 Excel 或 CSV 格式。'); }
  };
  r.readAsArrayBuffer(f);
}
function digest(data) {
  if (!(window.crypto && crypto.subtle)) return Promise.resolve('');
  return crypto.subtle.digest('SHA-256', data).then(function (b) {
    return Array.prototype.map.call(new Uint8Array(b), function (x) { return ('0' + x.toString(16)).slice(-2); }).join('').slice(0, 16);
  }).catch(function () { return ''; });
}
function openWorkbook(data, name) {
  wb = XLSX.read(data, { type: 'array', cellNF: true, cellDates: true });
  S.fileName = name;
  var ss = $('sheetSel'); ss.innerHTML = '';
  var best = 0, bestN = -1;
  wb.SheetNames.forEach(function (n, i) {
    ss.appendChild(el('option', { value: n, text: n }));
    var t = Core.sheetToTable(XLSX, wb.Sheets[n]); var m = Core.autoMap(t.headers, Y());
    var sc = Object.keys(m).length * 1000 + t.rows.length; if (m.id && sc > bestN) { bestN = sc; best = i; }
  });
  ss.value = wb.SheetNames[best];
  $('sheetRow').style.display = wb.SheetNames.length > 1 ? 'flex' : 'none';
  digest(data).then(function (h) { S.fileHash = h; useSheet(ss.value, true, true); });
}
$('sheetSel').addEventListener('change', function () { useSheet(this.value, true, false); });
function useSheet(n, remap, isNewFile) {
  var before = S.table ? (S.fileName + '／' + S.table.sheet) : '';
  S.table = Core.sheetToTable(XLSX, wb.Sheets[n]); S.table.sheet = n;
  if (remap) S.map = Core.autoMap(S.table.headers, Y());
  S.cfg.excluded = []; S.overrides = {};
  logChange('1 匯入資料', isNewFile ? '讀取試算表' : '切換工作表', before,
    S.fileName + '／' + n + '（' + S.table.rows.length + ' 列' + (S.fileHash ? '，SHA-256 ' + S.fileHash : '') + '）');
  var mapped = Object.keys(S.map).length;
  msg('impMsg', S.map.id ? 'ok' : 'warn', '已讀取「' + S.fileName + '」工作表「' + n + '」：' + S.table.rows.length + ' 列、' + S.table.headers.length + ' 欄，自動對應 ' + mapped + ' 個欄位。' + (S.map.id ? '' : '找不到「溪流編號」欄，請到「欄位對應」手動指定。'));
  recompute(); render(); save();
}
function msg(id, cls, text, link) {
  var box = $(id); box.innerHTML = '';
  var m = el('div', { class: 'msg ' + cls, text: text });
  if (link) { m.appendChild(document.createElement('br')); m.appendChild(el('a', { href: link, target: '_blank', rel: 'noopener', text: link })); }
  box.appendChild(m);
}

$('urlBtn').addEventListener('click', function () {
  var u = $('url').value.trim(); if (!u) { msg('urlMsg', 'err', '請先貼上分享連結。'); return; }
  var d = directUrl(u);
  msg('urlMsg', 'ok', '讀取中…');
  fetch(d, { redirect: 'follow' }).then(function (res) { if (!res.ok) throw new Error('HTTP ' + res.status); return res.arrayBuffer(); })
    .then(function (buf) { openWorkbook(new Uint8Array(buf), '雲端試算表（' + u + '）'); msg('urlMsg', 'ok', '已從連結讀取。'); })
    .catch(function () { msg('urlMsg', 'warn', '無法直接讀取此連結（可能未公開分享，或雲端服務不允許跨網站讀取）。請點下方直接下載連結取得檔案，再拖曳到左側匯入：', d); });
});
function directUrl(u) {
  var m = u.match(/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (m) { var g = u.match(/[#&?]gid=(\d+)/); return 'https://docs.google.com/spreadsheets/d/' + m[1] + '/export?format=xlsx' + (g ? '&gid=' + g[1] : ''); }
  var dm = u.match(/drive\.google\.com\/(?:file\/d\/|open\?id=)([a-zA-Z0-9_-]+)/);
  if (dm) return 'https://drive.google.com/uc?export=download&id=' + dm[1];
  if (/1drv\.ms|onedrive\.live\.com|sharepoint\.com/.test(u)) { var b = u.replace(/([?&])download=\d&?/, '$1').replace(/[?&]$/, ''); return b + (b.indexOf('?') >= 0 ? '&' : '?') + 'download=1'; }
  return u;
}

/* ---------- 計算 ---------- */
function recompute() {
  if (!S.table) { recs = []; sel = { list: [], perCat: {} }; return; }
  var cfg = S.cfg; cfg.th = { y15: +cfg.y15 || 15, y5: +cfg.y5 || 5, Y1: Y() - 1, Y2: Y() - 2 };
  recs = Core.buildRecords(S.table, S.map, { pctMode: cfg.pctMode });
  var have = {}; recs.forEach(function (r) { have[r.id] = 1; });
  cfg.customs.forEach(function (c) {
    c._missing = []; c._found = 0;
    Core.parseIds(c.ids).forEach(function (id) {
      if (have[id]) c._found++;
      else { c._missing.push(id); if (c.addMissing) { recs.push({ _i: 100000 + recs.length, id: id, stub: true }); have[id] = 1; } }
    });
  });
  Core.assign(recs, cfg);
  recs.forEach(function (r) { var o = S.overrides[r.id]; if (o) for (var k in o) if (/^(reasonOut|county|town|village)$/.test(k)) r[k] = o[k]; });
  Core.syncCats(recs, cfg);
  sel = Core.select(recs, cfg);
  if (!S.titleCustom) S.title = '新增及調整土石流潛勢溪流調查評估之' + sel.list.length + '條建議現勘清單';
  if (!S.noteCustom) S.note = Core.noteText(sel, cfg);
}
function update() { recompute(); render(); save(); }

/* ---------- 畫面 ---------- */
function render() {
  renderBar();
  $('subline').textContent = S.table ? (S.fileName + '・' + S.table.sheet + '・' + recs.length + ' 條候選') : '尚未匯入資料';
  $('logCount').textContent = S.log.length ? String(S.log.length) : '';
  ({ import: renderImport, map: renderMap, rules: renderRules, list: renderList, export: renderExport, log: renderLog })[S.view]();
}
function renderBar() {
  var bar = $('bar'); bar.innerHTML = '';
  var total = parseInt(S.cfg.total, 10), n = sel.list.length;
  var denom = total > 0 ? Math.max(total, n) : Math.max(n, 1), parts = [];
  S.cfg.cats.forEach(function (c) {
    var t = (sel.perCat[c.key] || {}).taken || 0; if (!t) return; parts.push(c.label + ' ' + t);
    var s = el('div', { class: 'seg' + (c.standby ? ' standby' : ''), title: c.label + '：' + t + ' 條', text: t / denom > 0.06 ? c.label + ' ' + t : String(t) });
    s.style.flexBasis = (t / denom * 100) + '%'; s.style.background = colorOf(c.key); bar.appendChild(s);
  });
  bar.setAttribute('aria-label', parts.length ? parts.join('、') : '尚無入選溪流');
  $('tallyN').textContent = n;
  $('tallyT').textContent = total > 0 ? '／' + total + ' 條' + (n < total ? '（尚缺 ' + (total - n) + '）' : '') : '／不限';
}
function renderImport() {
  var box = $('dataInfo');
  if (!S.table) { box.style.display = 'none'; return; }
  box.style.display = 'block'; box.innerHTML = '';
  var cnt = {}; recs.forEach(function (r) { cnt[r.catLabel] = (cnt[r.catLabel] || 0) + 1; });
  box.appendChild(el('h2', { text: '資料概況' }));
  var dl = el('dl', { class: 'kv' });
  function kv(k, v) { dl.appendChild(el('dt', { text: k })); dl.appendChild(el('dd', { text: v })); }
  kv('檔案', S.fileName + '（' + S.table.sheet + '，標題列第 ' + (S.table.headerRow + 1) + ' 列）');
  kv('候選溪流', recs.length + ' 條');
  kv('篩選類別', Object.keys(cnt).map(function (k) { return k + ' ' + cnt[k]; }).join('、') || '－');
  box.appendChild(dl);
  box.appendChild(el('div', { class: 'row', style: 'margin-top:14px' }, [el('button', { class: 'btn primary', text: '下一步：確認欄位對應', onclick: function () { show('map'); } })]));
}
function renderMap() {
  var tb = $('mapTbl').tBodies[0]; tb.innerHTML = '';
  if (!S.table) { tb.appendChild(el('tr', {}, [el('td', { colspan: 3, class: 'empty', text: '請先在「匯入資料」選擇試算表。' })])); return; }
  var first = S.table.rows[0] || {};
  Core.FIELDS(Y()).forEach(function (f) {
    var s = el('select', { 'aria-label': f.label });
    s.appendChild(el('option', { value: '', text: '（不對應）' }));
    S.table.headers.forEach(function (h) { var o = el('option', { value: h, text: h }); if (S.map[f.key] === h) o.selected = true; s.appendChild(o); });
    s.addEventListener('change', function () {
      logChange('2 欄位對應', f.label, S.map[f.key] || '（不對應）', s.value || '（不對應）');
      if (s.value) S.map[f.key] = s.value; else delete S.map[f.key]; update();
    });
    var v = S.map[f.key] != null ? first[S.map[f.key]] : '';
    tb.appendChild(el('tr', { class: f.req && !S.map[f.key] ? 'miss' : '' }, [
      el('td', {}, [f.label, f.req ? el('span', { class: 'req', text: ' *' }) : null, f.helper ? el('div', { class: 'muted', text: f.key === 'order' ? '作為類內排序依據' : '有值者自動判定為致災溪流' }) : null]),
      el('td', {}, [s]),
      el('td', {}, [el('span', { class: 'sample', text: v == null ? '' : (v instanceof Date ? v.toLocaleDateString() : String(v)) })])]));
  });
  $('pctMode').value = S.cfg.pctMode || 'auto';
}
function mapText(m) { return Object.keys(m).map(function (k) { return k + '=' + m[k]; }).join('；'); }
$('remap').addEventListener('click', function () {
  if (!S.table) return; var b = mapText(S.map); S.map = Core.autoMap(S.table.headers, Y());
  logChange('2 欄位對應', '重新自動對應', b, mapText(S.map)); update(); toast('已重新自動對應欄位');
});
$('pctMode').addEventListener('change', function () { var o = $('pctMode').options; logChange('2 欄位對應', '崩塌率數值格式', S.cfg.pctMode, this.value); S.cfg.pctMode = this.value; update(); });

function renderRules() {
  $('total').value = S.cfg.total; $('year').value = S.cfg.year; $('autoCls').checked = !!S.cfg.autoClassify; $('y15').value = S.cfg.y15; $('y5').value = S.cfg.y5;
  var tb = $('catTbl').tBodies[0]; tb.innerHTML = '';
  if (!S.cfg.cats.length) tb.appendChild(el('tr', {}, [el('td', { colspan: 8, class: 'empty', text: '匯入資料後，這裡會列出資料中的各篩選類別。' })]));
  S.cfg.cats.forEach(function (c, i) {
    var st = sel.perCat[c.key] || { avail: 0, taken: 0 };
    var isC = c.key.indexOf('c:') === 0;
    var L = '3 篩選設定／' + c.label;
    var chk = el('input', { type: 'checkbox', 'aria-label': '啟用' + c.label }); chk.checked = c.on;
    chk.addEventListener('change', function () { logChange(L, '啟用', c.on ? '是' : '否', chk.checked ? '是' : '否'); c.on = chk.checked; update(); });
    var q = el('input', { type: 'number', min: '0', placeholder: '不限', value: c.quota, 'aria-label': c.label + '上限數量' });
    q.addEventListener('change', function () { logChange(L, '上限數量', c.quota || '不限', q.value || '不限'); c.quota = q.value; update(); });
    var sb = el('select', { 'aria-label': c.label + '註記屬性' }, [el('option', { value: '0', text: '優先納入' }), el('option', { value: '1', text: '遞補' })]);
    sb.value = c.standby ? '1' : '0'; sb.style.width = 'auto';
    sb.addEventListener('change', function () { logChange(L, '註記屬性', c.standby ? '遞補' : '優先納入', sb.value === '1' ? '遞補' : '優先納入'); c.standby = sb.value === '1'; update(); });
    var sorts = el('div', { class: 'sortsel' });
    if (isC) sorts.appendChild(el('span', { class: 'muted', text: '依輸入順序' }));
    else for (var k = 0; k < 3; k++) (function (k) {
      var s = el('select', { 'aria-label': c.label + '排序' + (k + 1) });
      s.appendChild(el('option', { value: '', text: k ? '（無）' : '（試算表列序）' }));
      Object.keys(Core.SORT_KEYS).forEach(function (sk) { if (sk === 'row' || (sk === 'order' && !S.map.order)) return; s.appendChild(el('option', { value: sk, text: Core.SORT_KEYS[sk].label })); });
      s.value = c.sort[k] || '';
      s.addEventListener('change', function () { var b = sortText(c.sort); c.sort[k] = s.value; c.sort = c.sort.slice(0, 3); logChange(L, '類內排序依據', b, sortText(c.sort)); update(); });
      sorts.appendChild(s);
    })(k);
    var mv = el('span', { class: 'mv' }, [el('button', { text: '▲', 'aria-label': '上移', onclick: function () { move(i, i - 1); } }), el('button', { text: '▼', 'aria-label': '下移', onclick: function () { move(i, i + 1); } })]);
    var auto = recs.some(function (r) { return r.catKey === c.key && r.autoCat; });
    var tr = el('tr', { class: c.on ? '' : 'off', draggable: 'true' }, [
      el('td', {}, [el('span', { class: 'handle', text: '⠿', 'aria-hidden': 'true' }), mv, ' ' + (i + 1)]),
      el('td', {}, [chk]),
      el('td', {}, [el('span', { class: 'sw', style: 'background:' + colorOf(c.key) }), c.label,
        isC ? el('span', { class: 'pill', style: 'margin-left:6px', text: '業主指定' }) : null,
        auto ? el('span', { class: 'pill auto', style: 'margin-left:6px', text: '含自動判定' }) : null]),
      el('td', { text: String(c.count) }), el('td', {}, [q]), el('td', {}, [el('b', { text: String(st.taken) })]), el('td', {}, [sb]), el('td', {}, [sorts])]);
    tr.addEventListener('dragstart', function (e) { e.dataTransfer.setData('text/plain', String(i)); tr.classList.add('drag'); });
    tr.addEventListener('dragend', function () { tr.classList.remove('drag'); });
    tr.addEventListener('dragover', function (e) { e.preventDefault(); });
    tr.addEventListener('drop', function (e) { e.preventDefault(); var from = +e.dataTransfer.getData('text/plain'); if (!isNaN(from) && from !== i) move(from, i); });
    tb.appendChild(tr);
  });
  renderCustoms();
}
function orderText() { return S.cfg.cats.map(function (c) { return c.label; }).join(' > '); }
function move(from, to) {
  var c = S.cfg.cats; if (to < 0 || to >= c.length) return;
  var b = orderText(); var it = c.splice(from, 1)[0]; c.splice(to, 0, it);
  logChange('3 篩選設定', '類別優先順序', b, orderText()); update();
}
var FIELD_NAMES = { total: '調查數量', year: '調查年度', y15: '第1類門檻（年）', y5: '第3類門檻（年）' };
['total', 'year', 'y15', 'y5'].forEach(function (id) {
  $(id).addEventListener('change', function () {
    logChange('3 篩選設定', FIELD_NAMES[id], S.cfg[id], this.value); S.cfg[id] = this.value;
    if (id === 'year' && S.table) { var m = Core.autoMap(S.table.headers, Y()); ['ev1', 'ev2', 'red'].forEach(function (k) { if (m[k]) S.map[k] = m[k]; }); }
    update();
  });
});
$('autoCls').addEventListener('change', function () { logChange('3 篩選設定', '空白類別自動判定', S.cfg.autoClassify ? '是' : '否', this.checked ? '是' : '否'); S.cfg.autoClassify = this.checked; update(); });

function renderCustoms() {
  var box = $('customs'); box.innerHTML = '';
  S.cfg.customs.forEach(function (c, idx) {
    var L = '3 業主指定類別／' + c.name;
    var name = el('input', { type: 'text', value: c.name, placeholder: '例：重大災例、陳情案' });
    name.addEventListener('change', function () { var v = name.value.trim() || '業主指定'; logChange(L, '類別名稱', c.name, v); c.name = v; update(); });
    var reason = el('input', { type: 'text', value: c.reason || '', placeholder: '空白時沿用試算表原因' });
    reason.addEventListener('change', function () { var v = reason.value.trim(); logChange(L, '初步篩選辦理原因', c.reason, v); c.reason = v; update(); });
    var ids = el('textarea', { rows: '3', placeholder: '貼上溪流編號，以換行、逗號或空白分隔，例：投縣DF187, 花縣DF024' }); ids.value = c.ids || '';
    ids.addEventListener('change', function () {
      var a = Core.parseIds(c.ids), b = Core.parseIds(ids.value);
      var add = b.filter(function (x) { return a.indexOf(x) < 0; }), rem = a.filter(function (x) { return b.indexOf(x) < 0; });
      logChange(L, '溪流編號', a.join('、'), b.join('、') + (add.length || rem.length ? '（新增：' + (add.join('、') || '無') + '；移除：' + (rem.join('、') || '無') + '）' : ''));
      c.ids = ids.value; update();
    });
    var add = el('input', { type: 'checkbox' }); add.checked = !!c.addMissing;
    add.addEventListener('change', function () { logChange(L, '找不到的編號仍加入', c.addMissing ? '是' : '否', add.checked ? '是' : '否'); c.addMissing = add.checked; update(); });
    var n = Core.parseIds(c.ids).length;
    var stat = el('div', { class: 'muted', style: 'margin-top:6px' }, ['已輸入 ' + n + ' 條，資料中找到 ' + (c._found || 0) + ' 條', c._missing && c._missing.length ? el('span', { class: 'req', text: '；找不到：' + c._missing.join('、') }) : null]);
    box.appendChild(el('div', { class: 'custom', style: 'border-left-color:' + colorOf('c:' + c.uid) }, [
      el('div', { class: 'grid3' }, [
        el('div', {}, [el('label', { class: 'f', text: '類別名稱' }), name, el('label', { class: 'f', style: 'margin-top:8px', text: '初步篩選辦理原因' }), reason]),
        el('div', {}, [el('label', { class: 'f', text: '溪流編號' }), ids, stat])]),
      el('div', { class: 'row', style: 'margin-top:8px' }, [
        el('label', { class: 'row', style: 'gap:6px;font-size:.88rem' }, [add, '資料中找不到的編號仍加入名單（其餘欄位可於清冊中補填）']),
        el('button', { class: 'btn small ghost', style: 'margin-left:auto', text: '刪除此類別', onclick: function () {
          logChange('3 業主指定類別', '刪除類別', c.name + '（' + Core.parseIds(c.ids).join('、') + '）', '');
          S.cfg.customs.splice(idx, 1); S.cfg.cats = S.cfg.cats.filter(function (k) { return k.key !== 'c:' + c.uid; }); update();
        } })])
    ]));
  });
}
$('addCustom').addEventListener('click', function () {
  var uid = Date.now().toString(36);
  S.cfg.customs.push({ uid: uid, name: '業主指定', reason: '', ids: '', addMissing: false });
  logChange('3 業主指定類別', '新增類別', '', '業主指定'); update(); toast('已新增類別，請輸入溪流編號');
});

var FIELD_LABEL = { county: '縣市', town: '鄉鎮市區', village: '村里', reasonOut: '初步篩選辦理原因' };
function renderList() {
  $('pvTitle').textContent = S.title; $('pvNote').textContent = S.note;
  var cf = $('catFilter'), cur = cf.value; cf.innerHTML = ''; cf.appendChild(el('option', { value: '', text: '全部類別' }));
  S.cfg.cats.forEach(function (c) { if ((sel.perCat[c.key] || {}).taken) cf.appendChild(el('option', { value: c.key, text: c.label })); }); cf.value = cur;
  var wrap = $('listWrap'); wrap.innerHTML = '';
  if (!sel.list.length) { wrap.appendChild(el('div', { class: 'empty', text: S.table ? '目前設定沒有任何入選溪流，請到「篩選設定」檢查啟用類別與數量。' : '請先匯入試算表。' })); renderExcluded(); return; }
  var y1 = Y() - 1, y2 = Y() - 2;
  var t = el('table', { class: 'list' });
  var thead = el('thead');
  var h1 = el('tr'), h2 = el('tr');
  [['項次'], ['溪流編號'], ['縣市'], ['鄉鎮市區'], ['村里'], ['篩選類別'], ['初步篩選辦理原因'], ['事件後新增崩塌率(%)', 2], ['集水區崩塌率(%)'], ['風險等級'], [y1 + '年紅色警戒次數'], ['距前次調查年'], ['保全住戶數量'], ['移出']].forEach(function (h) {
    var th = el('th', { text: h[0] }); if (h[1]) th.setAttribute('colspan', '2'); else th.setAttribute('rowspan', '2'); h1.appendChild(th);
  });
  h2.appendChild(el('th', { text: y1 + '年' })); h2.appendChild(el('th', { text: y2 + '年' }));
  thead.appendChild(h1); thead.appendChild(h2); t.appendChild(thead);
  var tb = el('tbody'), q = $('q').value.trim(), catMap = {};
  S.cfg.cats.forEach(function (c) { catMap[c.key] = c; });
  var firstSb = -1; sel.list.some(function (r, i) { if (catMap[r.catKey] && catMap[r.catKey].standby) { firstSb = i; return true; } });
  sel.list.forEach(function (r, i) {
    if (cf.value && r.catKey !== cf.value) return;
    var cells = Core.rowCells(r, i);
    if (q && cells.join(' ').indexOf(q) < 0) return;
    var tr = el('tr', { class: i === firstSb && i > 0 ? 'sbstart' : '' });
    var editable = { 2: 'county', 3: 'town', 4: 'village', 6: 'reasonOut' };
    cells.forEach(function (v, k) {
      var td = k === 5 ? el('td', {}, [el('span', { class: 'chip', style: 'background:' + colorOf(r.catKey), text: v })]) : el('td', { class: k === 6 ? 'l' : '', text: v });
      var fk = editable[k];
      if (fk) {
        td.setAttribute('contenteditable', 'true'); td.setAttribute('aria-label', r.id + ' ' + FIELD_LABEL[fk]);
        if (S.overrides[r.id] && S.overrides[r.id][fk] != null) td.classList.add('edited');
        td.addEventListener('blur', function () {
          var nv = td.textContent.trim(); if (nv === v) return;
          logChange('4 篩選清冊／' + r.id, FIELD_LABEL[fk], v, nv);
          S.overrides[r.id] = S.overrides[r.id] || {}; S.overrides[r.id][fk] = nv; update();
        });
        td.addEventListener('keydown', function (e) { if (e.key === 'Enter') { e.preventDefault(); td.blur(); } });
      }
      tr.appendChild(td);
    });
    tr.appendChild(el('td', {}, [el('button', { class: 'x', text: '✕', 'aria-label': '移出 ' + r.id, title: '移出名單', onclick: function () {
      logChange('4 篩選清冊／' + r.id, '移出名單', '項次 ' + (i + 1) + '（' + r.catLabel + '）', '已移出');
      S.cfg.excluded.push(r.id); update(); toast('已移出 ' + r.id + '，後續溪流自動遞補');
    } })]));
    tb.appendChild(tr);
  });
  t.appendChild(tb); wrap.appendChild(t);
  renderExcluded();
}
function renderExcluded() {
  var ex = S.cfg.excluded; $('exSum').textContent = '已移出 ' + ex.length + ' 條';
  var box = $('exList'); box.innerHTML = '';
  ex.forEach(function (id) {
    box.appendChild(el('button', { class: 'btn small', text: '↺ ' + id, title: '放回名單', onclick: function () {
      logChange('4 篩選清冊／' + id, '放回名單', '已移出', '放回'); S.cfg.excluded = ex.filter(function (x) { return x !== id; }); update();
    } }));
  });
  $('exDetails').style.display = ex.length ? 'block' : 'none';
}
$('q').addEventListener('input', renderList); $('catFilter').addEventListener('change', renderList);

function renderExport() {
  $('title').value = S.title; $('note').value = S.note;
  var dl = $('expKv'); dl.innerHTML = '';
  function kv(k, v) { dl.appendChild(el('dt', { text: k })); dl.appendChild(el('dd', { text: v })); }
  kv('檔名', S.title + '.docx');
  kv('筆數', sel.list.length + ' 條');
  kv('各類別', S.cfg.cats.filter(function (c) { return (sel.perCat[c.key] || {}).taken; }).map(function (c) { return c.label + ' ' + sel.perCat[c.key].taken; }).join('、') || '－');
  kv('表頭年度', (Y() - 1) + '年、' + (Y() - 2) + '年');
  $('expBtn').disabled = !sel.list.length;
}
$('title').addEventListener('change', function () { var v = this.value.trim(); logChange('5 匯出 Word', '清冊標題', S.title, v); S.title = v; S.titleCustom = true; save(); renderExport(); });
$('note').addEventListener('change', function () { logChange('5 匯出 Word', '表下註記', S.note, this.value); S.note = this.value; S.noteCustom = true; save(); });
$('titleReset').addEventListener('click', function () { var b = S.title; S.titleCustom = false; recompute(); logChange('5 匯出 Word', '清冊標題（自動產生）', b, S.title); render(); save(); });
$('noteReset').addEventListener('click', function () { var b = S.note; S.noteCustom = false; recompute(); logChange('5 匯出 Word', '表下註記（自動產生）', b, S.note); render(); save(); toast('已依目前名單重新產生註記'); });

function download(blob, name) {
  var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click();
  setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
}
function safeName(s) { return String(s).replace(/[\\\/:*?"<>|]/g, '_'); }
$('expBtn').addEventListener('click', function () {
  var btn = this; btn.disabled = true; $('expMsg').textContent = '產生中…';
  var fname = safeName(S.title || '建議現勘清單') + '.docx';
  Core.buildDocx(JSZip, TPL.b64, TPL.row, sel.list, { title: S.title, Y: Y(), note: S.note }).then(function (blob) {
    download(blob, fname);
    logChange('5 匯出 Word', '匯出檔案', '', fname + '（' + sel.list.length + ' 條：' + sel.list.map(function (r) { return r.id; }).join('、') + '）');
    save(); $('expMsg').textContent = '已匯出 ' + fname;
  }).catch(function (e) { $('expMsg').textContent = '匯出失敗：' + e.message; }).then(function () { btn.disabled = false; });
});

/* ---------- 變更紀錄頁 ---------- */
function renderLog() {
  var wrap = $('logWrap'); wrap.innerHTML = '';
  var q = $('logQ').value.trim();
  var rows = S.log.map(function (e, i) { return [i + 1, e]; }).reverse().filter(function (x) {
    var e = x[1]; return !q || [e.loc, e.item, e.from, e.to].join(' ').indexOf(q) >= 0;
  });
  $('logSum').textContent = '共 ' + S.log.length + ' 筆' + (q ? '，符合 ' + rows.length + ' 筆' : '');
  if (!rows.length) { wrap.appendChild(el('div', { class: 'empty', text: S.log.length ? '沒有符合的紀錄。' : '尚無變更紀錄。匯入資料、調整設定或修改清冊後，每一項更動都會記錄在這裡。' })); return; }
  var t = el('table', { class: 'log' });
  t.appendChild(el('thead', {}, [el('tr', {}, ['#', '時間', '版本', '更動位置', '更動項目', '原值', '新值'].map(function (h) { return el('th', { text: h }); }))]));
  var tb = el('tbody');
  rows.slice(0, 1000).forEach(function (x) {
    var e = x[1];
    tb.appendChild(el('tr', {}, [el('td', { text: String(x[0]) }), el('td', { class: 'nw', text: fmtTime(e.t) }), el('td', { text: e.v }), el('td', { text: e.loc }), el('td', { text: e.item }),
      el('td', { class: 'from' }, [longText(e.from)]), el('td', { class: 'to' }, [longText(e.to)])]));
  });
  t.appendChild(tb); wrap.appendChild(t);
  if (rows.length > 1000) wrap.appendChild(el('p', { class: 'muted', text: '僅顯示最新 1000 筆，完整紀錄請匯出 CSV。' }));
}
function longText(v) {
  if (v.length <= 160) return document.createTextNode(v);
  return el('details', {}, [el('summary', { text: v.slice(0, 80) + '…（共 ' + v.length + ' 字）' }), el('div', { text: v })]);
}
function fmtTime(iso) { var d = new Date(iso); function p(n) { return ('0' + n).slice(-2); } return (d.getFullYear() - 1911) + '/' + p(d.getMonth() + 1) + '/' + p(d.getDate()) + ' ' + p(d.getHours()) + ':' + p(d.getMinutes()) + ':' + p(d.getSeconds()); }
$('logQ').addEventListener('input', renderLog);
function stamp() { var d = new Date(); function p(n) { return ('0' + n).slice(-2); } return d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + '_' + p(d.getHours()) + p(d.getMinutes()); }
$('logCsv').addEventListener('click', function () {
  function c(v) { v = String(v == null ? '' : v); return /[",\n\r]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; }
  var lines = [['序號', '時間', '版本', '更動位置', '更動項目', '原值', '新值'].join(',')];
  S.log.forEach(function (e, i) { lines.push([i + 1, fmtTime(e.t), e.v, e.loc, e.item, e.from, e.to].map(c).join(',')); });
  download(new Blob(['\ufeff' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' }), '變更紀錄_' + stamp() + '.csv');
});
$('logClear').addEventListener('click', function () {
  if (!confirm('確定清除全部 ' + S.log.length + ' 筆變更紀錄？建議先匯出 CSV 或設定檔留存。')) return;
  var n = S.log.length; S.log = []; logChange('變更紀錄', '清除紀錄', n + ' 筆', '0 筆'); save(); render();
});

/* ---------- 設定檔匯出／匯入（可提交到 GitHub 留存版本） ---------- */
$('cfgExport').addEventListener('click', function () {
  var out = { app: '土石流潛勢溪流現勘名單篩選', version: VERSION, savedAt: new Date().toISOString(),
    source: S.table ? { file: S.fileName, sheet: S.table.sheet, rows: S.table.rows.length, sha256: S.fileHash || '' } : null,
    map: S.map, cfg: stripRuntime(S.cfg), overrides: S.overrides,
    title: S.title, titleCustom: S.titleCustom, note: S.note, noteCustom: S.noteCustom,
    selected: sel.list.map(function (r, i) { return { no: i + 1, id: r.id, cat: r.catLabel }; }), log: S.log };
  logChange('設定檔', '匯出設定檔', '', '篩選設定_' + stamp() + '.json');
  out.log = S.log; save(); render();
  download(new Blob([JSON.stringify(out, null, 2)], { type: 'application/json' }), '篩選設定_' + stamp() + '.json');
});
function stripRuntime(cfg) { var c = clone(cfg); delete c.th; (c.customs || []).forEach(function (x) { delete x._missing; delete x._found; }); return c; }
$('cfgImport').addEventListener('click', function () { $('cfgFile').click(); });
$('cfgFile').addEventListener('change', function () {
  var f = this.files[0]; this.value = ''; if (!f) return;
  var r = new FileReader();
  r.onload = function () {
    try {
      var o = JSON.parse(r.result); if (!o.cfg) throw new Error('檔案中沒有 cfg 設定');
      S.cfg = Object.assign(clone(DEF_CFG), o.cfg); S.map = o.map || S.map; S.overrides = o.overrides || {};
      S.title = o.title || ''; S.titleCustom = !!o.titleCustom; S.note = o.note || ''; S.noteCustom = !!o.noteCustom;
      var merged = (o.log || []).slice(); var seen = {}; merged.forEach(function (e) { seen[e.t + e.item + e.to] = 1; });
      S.log.forEach(function (e) { if (!seen[e.t + e.item + e.to]) merged.push(e); });
      merged.sort(function (a, b) { return a.t < b.t ? -1 : a.t > b.t ? 1 : 0; }); S.log = merged;
      logChange('設定檔', '匯入設定檔', '', f.name + '（版本 ' + (o.version || '?') + '，' + (o.savedAt || '') + '）');
      update(); toast('已匯入設定' + (S.table ? '' : '，請再匯入試算表'));
    } catch (e) { toast('無法匯入設定檔：' + e.message); }
  };
  r.readAsText(f);
});

/* ---------- 啟動 ---------- */
$('ver').textContent = 'v' + VERSION;
load();
if (S.table) { try { recompute(); } catch (e) { S.table = null; } }
var h = location.hash.slice(1);
show(/^(import|map|rules|list|export|log)$/.test(h) ? h : (S.view || 'import'));
})();
