/* ===== 核心邏輯：欄位對應、分類、篩選、Word 匯出 ===== */
var Core = (function () {
  var DASH = '－';

  function FIELDS(Y) {
    var y1 = Y - 1, y2 = Y - 2;
    return [
      { key: 'id', label: '溪流編號', req: true, syn: ['溪流編號', '土石流編號', '編號', 'DF編號', '溪流代碼'] },
      { key: 'county', label: '縣市', syn: ['縣市', '縣市別', '縣市名稱'] },
      { key: 'town', label: '鄉鎮市區', syn: ['鄉鎮市區', '鄉鎮', '鄉鎮市', '鄉鎮區', '鄉鎮名稱'] },
      { key: 'village', label: '村里', syn: ['村里', '村里名稱', '里別'] },
      { key: 'cat', label: '篩選類別', syn: ['篩選類別', '類別', '優先類別', '調查類別'] },
      { key: 'reason', label: '初步篩選辦理原因', syn: ['初步篩選辦理原因', '辦理原因', '篩選原因', '原因'] },
      { key: 'ev1', label: y1 + '年事件後新增崩塌率', type: 'pct', syn: [y1 + '年事件崩塌率', y1 + '年新增崩塌率', y1 + '年崩塌率', y1 + '年'] },
      { key: 'ev2', label: y2 + '年事件後新增崩塌率', type: 'pct', syn: [y2 + '年事件崩塌率', y2 + '年新增崩塌率', y2 + '年崩塌率', y2 + '年'] },
      { key: 'wsr', label: '集水區崩塌率', type: 'pct', syn: ['集水區崩塌率', '集水區崩塌率%', '崩塌率'] },
      { key: 'risk', label: '風險等級', syn: ['風險等級', '風險'] },
      { key: 'red', label: y1 + '年紅色警戒次數', type: 'num', syn: [y1 + '年紅色警戒次數', '紅色警戒次數', '紅警次數', '紅色警戒'] },
      { key: 'yrs', label: '距前次調查年', type: 'num', syn: ['距前次調查年', '距前次調查年數', '距前次更新年', '距前次更新'] },
      { key: 'house', label: '保全住戶數量', type: 'num', syn: ['保全住戶數量', '保全住戶戶數', '保全住戶數', '保全住戶', '保全戶數'] },
      { key: 'order', label: '原始排序（如發文排序）', helper: true, type: 'num', syn: ['發文排序', '排序', '建議排序'] },
      { key: 'event', label: y1 + '年致災事件名稱', helper: true, syn: ['事件名稱', '致災事件', '災害事件'] }
    ];
  }

  function normH(s) {
    return String(s == null ? '' : s).replace(/[\s\r\n()（）%％]/g, '').replace(/[臺]/g, '台');
  }

  function autoMap(headers, Y) {
    var map = {}, used = {};
    var F = FIELDS(Y);
    var nh = headers.map(normH);
    // pass 1: exact, pass 2: contains
    [0, 1].forEach(function (pass) {
      F.forEach(function (f) {
        if (map[f.key] != null) return;
        for (var s = 0; s < f.syn.length; s++) {
          var target = normH(f.syn[s]);
          for (var i = 0; i < nh.length; i++) {
            if (used[i] || !nh[i]) continue;
            var ok = pass === 0 ? nh[i] === target : (target.length >= 3 && nh[i].indexOf(target) >= 0);
            if (ok) { map[f.key] = headers[i]; used[i] = 1; return; }
          }
        }
      });
    });
    return map;
  }

  /* ---- 讀取工作表 ---- */
  function detectHeaderRow(aoa) {
    var best = 0, bestScore = -1;
    var keys = ['溪流編號', '縣市', '鄉鎮', '村里', '風險等級', '保全住戶', '篩選類別'];
    for (var r = 0; r < Math.min(aoa.length, 15); r++) {
      var row = aoa[r] || [], sc = 0;
      row.forEach(function (c) { var n = normH(c); keys.forEach(function (k) { if (n.indexOf(k) >= 0) sc++; }); });
      if (sc > bestScore) { bestScore = sc; best = r; }
    }
    return best;
  }

  function sheetToTable(XLSX, ws) {
    var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
    var hr = detectHeaderRow(aoa);
    var headers = (aoa[hr] || []).map(function (h, i) { return h == null || h === '' ? '(欄' + XLSX.utils.encode_col(i) + ')' : String(h).trim(); });
    // percent format detection per column
    var pctCols = {};
    var range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    for (var c = range.s.c; c <= range.e.c; c++) {
      for (var r = range.s.r + hr + 1; r <= Math.min(range.e.r, range.s.r + hr + 400); r++) {
        var cell = ws[XLSX.utils.encode_cell({ r: r, c: c })];
        if (cell && cell.t === 'n' && cell.z && String(cell.z).indexOf('%') >= 0) { pctCols[headers[c - range.s.c]] = 1; break; }
      }
    }
    var rows = [];
    for (var i = hr + 1; i < aoa.length; i++) {
      var a = aoa[i]; if (!a) continue;
      var o = {}, any = false;
      headers.forEach(function (h, j) { var v = a[j]; if (v != null && v !== '') any = true; o[h] = v; });
      if (any) rows.push(o);
    }
    return { headers: headers, rows: rows, pctCols: pctCols, headerRow: hr };
  }

  /* ---- 格式 ---- */
  function isBlank(v) { return v == null || String(v).trim() === '' || String(v).trim() === DASH || String(v).trim() === '-'; }
  function num(v) {
    if (isBlank(v)) return null;
    if (typeof v === 'number') return v;
    var s = String(v).replace(/[,%％\s]/g, '');
    var n = parseFloat(s); return isNaN(n) ? null : n;
  }
  // 回傳百分比數值（0.15 代表 0.15%）
  function pctValue(v, isFraction) {
    if (isBlank(v)) return null;
    if (typeof v === 'string' && /[%％]/.test(v)) return num(v);
    var n = num(v); if (n == null) return null;
    return isFraction ? n * 100 : n;
  }
  function fmtPct(p) {
    if (p == null) return DASH;
    if (p === 0) return '0.00%';
    var d = 2;
    while (Math.abs(p) < Math.pow(10, -d) && d < 6) d++;
    var s = p.toFixed(d);
    return s + '%';
  }
  function fmtNum(v) {
    if (isBlank(v)) return DASH;
    if (typeof v === 'number') return String(Math.round(v * 1000) / 1000);
    return String(v).trim();
  }
  function cleanCat(s) {
    s = String(s == null ? '' : s).trim();
    return s.replace(/^[0-9０-９]+(?=[^0-9０-９類])/, '').replace(/^[_\-\.\s、]+/, '');
  }
  function txt(v) { return isBlank(v) ? DASH : String(v).trim(); }

  var RISK_RANK = { '高': 4, '中': 3, '低': 2, '持續觀察': 1 };
  function riskRank(v) { return RISK_RANK[String(v == null ? '' : v).trim()] || 0; }

  /* ---- 建立紀錄 ---- */
  function buildRecords(table, map, opts) {
    var pc = table.pctCols || {};
    function isFrac(key) {
      var h = map[key]; if (!h) return true;
      if (pc[h]) return true;
      if (opts && opts.pctMode === 'percent') return false;
      if (opts && opts.pctMode === 'fraction') return true;
      // auto：有任何數值 > 1 則視為百分比數值
      var mx = 0; table.rows.forEach(function (r) { var n = typeof r[h] === 'number' ? r[h] : null; if (n != null && n > mx) mx = n; });
      return mx <= 1;
    }
    var fr = { ev1: isFrac('ev1'), ev2: isFrac('ev2'), wsr: isFrac('wsr') };
    function g(r, k) { return map[k] ? r[map[k]] : null; }
    var out = [];
    table.rows.forEach(function (r, i) {
      var id = g(r, 'id'); if (isBlank(id)) return;
      out.push({
        _i: i, id: String(id).trim(),
        county: g(r, 'county'), town: g(r, 'town'), village: g(r, 'village'),
        rawCat: g(r, 'cat') == null ? '' : String(g(r, 'cat')).trim(),
        reason: g(r, 'reason'),
        ev1: pctValue(g(r, 'ev1'), fr.ev1), ev2: pctValue(g(r, 'ev2'), fr.ev2), wsr: pctValue(g(r, 'wsr'), fr.wsr),
        risk: g(r, 'risk'), red: g(r, 'red'), yrs: g(r, 'yrs'), house: g(r, 'house'),
        order: num(g(r, 'order')), event: g(r, 'event')
      });
    });
    return out;
  }

  /* 依流程圖自動判定（僅用於篩選類別空白者） */
  function autoClassify(rec, th) {
    var yrs = num(rec.yrs), house = num(rec.house), red = num(rec.red);
    if (yrs != null && yrs >= th.y15 && house != null && house > 0) return ['第1類', '距前次更新已逾' + th.y15 + '年且具保全住戶'];
    if (!isBlank(rec.event)) return ['第2類', th.Y1 + '年度颱風豪雨事件致災'];
    if (rec.ev1 != null && rec.ev1 > 0) return ['第2類', th.Y1 + '年度集水區具新生崩塌地'];
    if (rec.ev2 != null && rec.ev2 > 0) return ['第2類', th.Y2 + '年度集水區具新生崩塌地'];
    if (red != null && red > 0 && yrs != null && yrs >= th.y5) return ['第3類', th.Y1 + '年度發布紅色警戒且具前次更新滿' + th.y5 + '年'];
    return ['第4類', '依風險等級及前次更新年度排序'];
  }

  /* ---- 排序鍵 ---- */
  var SORT_KEYS = {
    order: { label: '原始排序（小→大）', fn: function (r) { return r.order == null ? Infinity : r.order; }, dir: 1 },
    risk: { label: '風險等級（高→低）', fn: function (r) { return riskRank(r.risk); }, dir: -1 },
    house: { label: '保全住戶數（多→少）', fn: function (r) { var n = num(r.house); return n == null ? -1 : n; }, dir: -1 },
    yrs: { label: '距前次調查年（久→近）', fn: function (r) { var n = num(r.yrs); return n == null ? -1 : n; }, dir: -1 },
    red: { label: '紅色警戒次數（多→少）', fn: function (r) { var n = num(r.red); return n == null ? -1 : n; }, dir: -1 },
    wsr: { label: '集水區崩塌率（高→低）', fn: function (r) { return r.wsr == null ? -1 : r.wsr; }, dir: -1 },
    ev1: { label: '事件後新增崩塌率（高→低）', fn: function (r) { return Math.max(r.ev1 == null ? -1 : r.ev1, r.ev2 == null ? -1 : r.ev2); }, dir: -1 },
    row: { label: '試算表列序', fn: function (r) { return r._i; }, dir: 1 }
  };
  function sorter(keys) {
    keys = (keys || []).filter(function (k) { return SORT_KEYS[k]; });
    return function (a, b) {
      for (var i = 0; i < keys.length; i++) {
        var K = SORT_KEYS[keys[i]], x = K.fn(a), y = K.fn(b);
        if (x !== y) return (x < y ? -1 : 1) * K.dir;
      }
      return a._i - b._i;
    };
  }

  function defaultRank(label) {
    var s = String(label);
    if (/地方政府/.test(s)) return 10;
    if (/主動提報|農村水保署/.test(s)) return 11;
    var t = s.match(/Type\s*([ABC])/i); if (t) return 20 + 'ABC'.indexOf(t[1].toUpperCase());
    var m = s.match(/第\s*([0-9０-９]+)\s*類/); if (m) return 30 + parseInt(m[1].replace(/[０-９]/g, function (c) { return c.charCodeAt(0) - 65296; }), 10);
    return 90;
  }
  function defaultSort(label, hasOrder) {
    if (hasOrder) return ['order', 'risk', 'house'];
    if (/第\s*1\s*類/.test(label)) return ['risk', 'house'];
    if (/第\s*2\s*類/.test(label)) return ['risk', 'house'];
    if (/第\s*3\s*類/.test(label)) return ['risk', 'yrs', 'house'];
    if (/第\s*4\s*類/.test(label)) return ['risk', 'yrs', 'red', 'house'];
    return ['risk', 'house'];
  }
  function defaultStandby(label) { return /第\s*[34]\s*類/.test(label); }

  /* 指派類別：自訂（業主指定）> 資料欄位 > 自動判定 */
  function assign(records, cfg) {
    var custom = {};
    (cfg.customs || []).forEach(function (c) {
      parseIds(c.ids).forEach(function (id) { if (!custom[id]) custom[id] = c; });
    });
    records.forEach(function (r) {
      var c = custom[r.id];
      if (c) { r.catKey = 'c:' + c.uid; r.catLabel = c.name; r.reasonOut = c.reason || txt(r.reason); return; }
      var label = cleanCat(r.rawCat), reason = r.reason;
      if (!label && cfg.autoClassify) {
        var a = autoClassify(r, cfg.th); label = a[0]; if (isBlank(reason)) reason = a[1]; r.autoCat = true;
      }
      if (!label) label = '未分類';
      r.catKey = 'd:' + label; r.catLabel = label; r.reasonOut = txt(reason);
    });
    return records;
  }

  function parseIds(s) {
    return String(s || '').split(/[\s,，、;；]+/).map(function (x) { return x.trim(); }).filter(Boolean);
  }

  /* 同步類別設定清單（保留使用者的順序/數量） */
  function syncCats(records, cfg) {
    var counts = {}, labels = {};
    records.forEach(function (r) { counts[r.catKey] = (counts[r.catKey] || 0) + 1; labels[r.catKey] = r.catLabel; });
    var hasOrder = records.some(function (r) { return r.order != null; });
    var existing = {}; (cfg.cats || []).forEach(function (c) { existing[c.key] = c; });
    var list = [];
    (cfg.cats || []).forEach(function (c) {
      var live = counts[c.key] != null || c.key.indexOf('c:') === 0 && (cfg.customs || []).some(function (x) { return 'c:' + x.uid === c.key; });
      if (live) { var cu = (cfg.customs || []).filter(function (x) { return 'c:' + x.uid === c.key; })[0]; c.label = cu ? cu.name : (labels[c.key] || c.label); c.count = counts[c.key] || 0; list.push(c); }
    });
    Object.keys(counts).forEach(function (k) {
      if (existing[k]) return;
      var lab = labels[k];
      var isC = k.indexOf('c:') === 0;
      var c = { key: k, label: lab, count: counts[k], quota: '', on: true, standby: isC ? false : defaultStandby(lab), sort: defaultSort(lab, hasOrder), rank: isC ? 0 : defaultRank(lab) };
      // 插入：依預設 rank 放到合適位置
      var pos = list.length;
      for (var i = 0; i < list.length; i++) { if ((list[i].rank == null ? 99 : list[i].rank) > c.rank) { pos = i; break; } }
      list.splice(pos, 0, c);
    });
    (cfg.customs || []).forEach(function (x) {
      var k = 'c:' + x.uid;
      if (!list.some(function (c) { return c.key === k; })) list.unshift({ key: k, label: x.name, count: 0, quota: '', on: true, standby: false, sort: ['row'], rank: 0 });
    });
    cfg.cats = list;
    return list;
  }

  function select(records, cfg) {
    var total = parseInt(cfg.total, 10); if (!(total > 0)) total = Infinity;
    var excluded = {}; (cfg.excluded || []).forEach(function (id) { excluded[id] = 1; });
    var taken = {}, out = [], perCat = {};
    var custOrder = {};
    (cfg.customs || []).forEach(function (c) { parseIds(c.ids).forEach(function (id, i) { if (custOrder[id] == null) custOrder[id] = i; }); });
    cfg.cats.forEach(function (c) {
      perCat[c.key] = { avail: 0, taken: 0 };
      if (!c.on) { perCat[c.key].avail = records.filter(function (r) { return r.catKey === c.key; }).length; return; }
      var pool = records.filter(function (r) { return r.catKey === c.key && !excluded[r.id] && !taken[r.id]; });
      perCat[c.key].avail = pool.length;
      if (c.key.indexOf('c:') === 0) pool.sort(function (a, b) { return custOrder[a.id] - custOrder[b.id]; });
      else pool.sort(sorter(c.sort));
      var q = parseInt(c.quota, 10); if (!(q >= 0)) q = Infinity;
      for (var i = 0; i < pool.length && perCat[c.key].taken < q && out.length < total; i++) {
        taken[pool[i].id] = 1; out.push(pool[i]); perCat[c.key].taken++;
      }
    });
    return { list: out, perCat: perCat };
  }

  /* 註解文字 */
  function catGroupText(labels) {
    var named = [], nums = [];
    labels.forEach(function (l) { var m = String(l).match(/^第\s*(\d+)\s*類$/); if (m) nums.push(+m[1]); else named.push(l); });
    nums.sort(function (a, b) { return a - b; });
    var segs = [], i = 0;
    while (i < nums.length) {
      var j = i; while (j + 1 < nums.length && nums[j + 1] === nums[j] + 1) j++;
      segs.push(i === j ? '第' + nums[i] + '類' : '第' + nums[i] + '~' + nums[j] + '類'); i = j + 1;
    }
    var all = named.concat(segs);
    if (all.length <= 1) return all.join('');
    return all.slice(0, -1).join('、') + '及' + all[all.length - 1];
  }
  function noteText(sel, cfg) {
    var pri = [], sb = [], priL = [], sbL = [];
    var catMap = {}; cfg.cats.forEach(function (c) { catMap[c.key] = c; });
    sel.list.forEach(function (r, i) { var c = catMap[r.catKey]; if (c && c.standby) { sb.push(i + 1); if (sbL.indexOf(c.label) < 0) sbL.push(c.label); } else { pri.push(i + 1); if (priL.indexOf(c.label) < 0) priL.push(c.label); } });
    if (!sel.list.length) return '';
    function rng(a) { return a.length ? (a[0] === a[a.length - 1] ? '第' + a[0] + '項' : '第' + a[0] + '~' + a[a.length - 1] + '項') : ''; }
    var s = '註：';
    if (pri.length) s += '項次' + rng(pri) + '(篩選類別：' + catGroupText(priL) + ')建議優先納入本年度現勘名單，未來若有地方政府或農村水保署主動提報新增及調整之土石流潛勢溪流，亦納入優先現勘名單';
    if (sb.length) s += (pri.length ? '，若仍有剩餘額度時，則以' : '') + '項次' + rng(sb) + '(' + catGroupText(sbL) + ')依序列入現勘名單';
    return s + '。';
  }

  function rowCells(r, i) {
    return [String(i + 1), txt(r.id), txt(r.county), txt(r.town), txt(r.village), r.catLabel || DASH, r.reasonOut || DASH,
      fmtPct(r.ev1), fmtPct(r.ev2), fmtPct(r.wsr), txt(r.risk), fmtNum(r.red), fmtNum(r.yrs), fmtNum(r.house)];
  }

  function esc(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }

  function buildDocx(JSZip, tplB64, rowTpl, list, opts) {
    return JSZip.loadAsync(tplB64, { base64: true }).then(function (zip) {
      return zip.file('word/document.xml').async('string').then(function (doc) {
        var rows = list.map(function (r, i) {
          var cells = rowCells(r, i), s = rowTpl;
          for (var k = 0; k < 14; k++) s = s.split('{{C' + k + '}}').join(esc(cells[k]));
          return s;
        }).join('');
        function R(a, b) { doc = doc.split(a).join(b); }
        R('{{TITLE}}', esc(opts.title)); R('{{Y1}}', String(opts.Y - 1)); R('{{Y2}}', String(opts.Y - 2)); R('{{NOTE}}', esc(opts.note || '')); R('{{ROWS}}', rows);

        zip.file('word/document.xml', doc);
        return zip.generateAsync({ type: opts.nodeBuffer ? 'nodebuffer' : 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', compression: 'DEFLATE' });
      });
    });
  }

  return {
    FIELDS: FIELDS, autoMap: autoMap, sheetToTable: sheetToTable, buildRecords: buildRecords, assign: assign, syncCats: syncCats,
    select: select, noteText: noteText, rowCells: rowCells, buildDocx: buildDocx, SORT_KEYS: SORT_KEYS, parseIds: parseIds,
    cleanCat: cleanCat, fmtPct: fmtPct, txt: txt, catGroupText: catGroupText
  };
})();
if (typeof module !== 'undefined') module.exports = Core;
