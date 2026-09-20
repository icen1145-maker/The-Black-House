/**
 * 看房板 · 数据后端（和时间表共用同一个 Web App URL）
 *
 * 读：GET  <exec URL>?view=houses
 * 写：POST <exec URL>  body = { view:'houses', ops:[...] }
 *   {type:'update', addr:'原地址', fields:{status:'Sold', cons:'...'}}
 *   {type:'add',    fields:{addr:'...', price:'1.1m', ...}}
 *
 *   {type:'fav',    addr:'...', on:true}          ← 爱心收藏
 *   {type:'no',     addr:'...', on:true}          ← 标记「不要了」
 *   {type:'attr',    addr:'...', fields:{bd:3, ba:2, sqft:1400, ...}}  ← 手填/粘贴解析出来的硬数据
 *   {type:'attrEnv', addr:'...', fields:{sf, noise, walk, fs, face, at, src}}  ← 网页抓来的环境信息
 *
 * 只读写 2026 tab 里按表头名找到的 7 列，不增删列、不删行。
 * 收藏 / 不要 存在脚本属性里；房源属性写在单独的「房源属性」tab（见 Attrs.gs）。
 * 坐标用 Apps Script 自带的 Maps 地理编码，结果缓存在脚本属性里（不写进表格）。
 */

var HOUSE_SHEET = '2026';
var HOUSE_COLS = {            // 字段 → 表头（大小写、空格不敏感）
  addr:   ['地址', 'address'],
  cons:   ['硬伤', 'cons'],
  pros:   ['优点', 'pros'],
  price:  ['价格', 'price', 'list'],
  sold:   ['售价', 'sold', 'sold price'],
  status: ['状态', 'status'],
  oh:     ['open house', 'openhouse', 'oh']
};
var GEO_NEW_PER_CALL = 12;    // 每次请求最多新定位几处，避免超时

function houseSheet_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOUSE_SHEET);
  if (!sh) throw new Error('找不到工作表 ' + HOUSE_SHEET);
  return sh;
}

function normHead_(s) { return String(s || '').toLowerCase().replace(/\s+/g, ' ').trim(); }
function normAddr_(s) { return String(s || '').toLowerCase().replace(/[.,#]/g, ' ').replace(/\s+/g, ' ').trim(); }
function favKey_(addr) { return 'fav:' + normAddr_(addr).slice(0, 200); }
function noKey_(addr) { return 'no:' + normAddr_(addr).slice(0, 200); }

/** 表头 → {field: 列号(1-based)} */
function houseColMap_(sh) {
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var head = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(normHead_);
  var map = {};
  Object.keys(HOUSE_COLS).forEach(function (f) {
    for (var i = 0; i < head.length; i++) {
      if (HOUSE_COLS[f].indexOf(head[i]) >= 0) { map[f] = i + 1; return; }
    }
  });
  if (!map.addr) throw new Error('2026 tab 第一行没找到「地址」列');
  return map;
}

function houseRows_(sh, map) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  var width = Math.max.apply(null, Object.keys(map).map(function (k) { return map[k]; }));
  var vals = sh.getRange(2, 1, last - 1, width).getDisplayValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i], h = { r: i + 2 };
    Object.keys(map).forEach(function (f) { h[f] = String(r[map[f] - 1] || '').trim(); });
    if (h.addr) out.push(h);
  }
  return out;
}

/* ---------- 地理编码（缓存在脚本属性） ---------- */
function geoLookup_(addr, budget, cache, props) {
  var key = 'geo:' + normAddr_(addr).slice(0, 200);
  var hit = cache[key];
  if (hit) {
    try {
      var g = JSON.parse(hit);
      if (g.lat != null) return g;
      if (Date.now() - (g.t || 0) < 864e5) return null;      // 失败的一天内不重试
    } catch (e) {}
  }
  if (budget.left <= 0) return null;
  budget.left--;
  try {
    var q = /,\s*ca\b|california/i.test(addr) ? addr : addr + ', CA';
    var res = Maps.newGeocoder().setBounds(36.9, -122.6, 37.9, -121.4).geocode(q);
    if (res.status === 'OK' && res.results && res.results.length) {
      var top = res.results[0], loc = top.geometry.location;
      var t = top.geometry.location_type;
      var out = { lat: +loc.lat.toFixed(6), lng: +loc.lng.toFixed(6),
                  q: (t === 'ROOFTOP' || t === 'RANGE_INTERPOLATED') && !top.partial_match ? 'exact' : 'approx' };
      props.setProperty(key, JSON.stringify(out));
      return out;
    }
    props.setProperty(key, JSON.stringify({ t: Date.now() }));
  } catch (err) {
    budget.err = String(err);
  }
  return null;
}

function housesRead_() {
  var sh = houseSheet_();
  var map = houseColMap_(sh);
  var rows = houseRows_(sh, map);
  var budget = { left: GEO_NEW_PER_CALL };
  var props = PropertiesService.getScriptProperties();
  var cache = props.getProperties();          // 一次读完，避免每处房源各读一次
  var attrs = attrAll_();
  rows.forEach(function (h) {
    var g = geoLookup_(h.addr, budget, cache, props);
    if (g) { h.lat = g.lat; h.lng = g.lng; h.geo = g.q; }
    if (cache[favKey_(h.addr)] === '1') h.fav = true;
    if (cache[noKey_(h.addr)] === '1') h.no = true;
    var a = attrs[normAddr_(h.addr)];
    if (a) { var o = {}; ATTR_KEYS.forEach(function (k) { if (k !== 'addr' && a[k]) o[k] = a[k]; }); h.a = o; }
  });
  var res = { ok: true, sheet: HOUSE_SHEET, fields: Object.keys(map), houses: rows,
              at: new Date().toISOString() };
  if (budget.err) res.geoError = budget.err;
  return res;
}

function writeCell_(sh, row, col, v) {
  sh.getRange(row, col).setNumberFormat('@').setValue(String(v == null ? '' : v));
}

/** 批量写。只认 update / add；update 找不到地址会报错，不会乱写。 */
function housesWrite_(body) {
  var sh = houseSheet_();
  var map = houseColMap_(sh);
  var ops = body.ops || [];
  var props = PropertiesService.getScriptProperties();
  var errors = [];

  ops.forEach(function (op, idx) {
    var fields = op.fields || {};
    var rows = houseRows_(sh, map);
    var findRow = function (a) {
      var k = normAddr_(a);
      for (var i = 0; i < rows.length; i++) if (normAddr_(rows[i].addr) === k) return rows[i].r;
      return 0;
    };

    var row = 0;
    if (op.type === 'fav' || op.type === 'no') {   // 收藏 / 不要：只动脚本属性，两者互斥
      if (!String(op.addr || '').trim()) { errors.push('第 ' + (idx + 1) + ' 条：缺地址'); return; }
      var mine = op.type === 'fav' ? favKey_(op.addr) : noKey_(op.addr);
      var other = op.type === 'fav' ? noKey_(op.addr) : favKey_(op.addr);
      if (op.on) { props.setProperty(mine, '1'); props.deleteProperty(other); }
      else props.deleteProperty(mine);
      return;
    }
    if (op.type === 'attr' || op.type === 'attrEnv') {    // 写「房源属性」tab
      if (!String(op.addr || '').trim()) { errors.push('第 ' + (idx + 1) + ' 条：缺地址'); return; }
      var allow = op.type === 'attr' ? FACT_KEYS : ENV_KEYS;
      var keep = {};
      allow.forEach(function (k) { if (fields[k] !== undefined) keep[k] = fields[k]; });
      if (op.type === 'attrEnv' && keep.elec !== undefined) {       // 自动猜的电线不覆盖手选的
        var had = attrAll_()[normAddr_(op.addr)];
        if (had && had.elec) delete keep.elec;
      }
      attrSave_(op.addr, keep);
      return;
    }
    if (op.type === 'update') {
      row = findRow(op.addr);
      if (!row) { errors.push('第 ' + (idx + 1) + ' 条：表格里找不到 ' + op.addr); return; }
      if (fields.addr && normAddr_(fields.addr) !== normAddr_(op.addr)) {   // 改了地址，收藏/不要/属性跟着搬家
        [[favKey_(op.addr), favKey_(fields.addr)], [noKey_(op.addr), noKey_(fields.addr)]].forEach(function (pair) {
          var v = props.getProperty(pair[0]);
          if (v) { props.setProperty(pair[1], v); props.deleteProperty(pair[0]); }
        });
        attrRename_(op.addr, fields.addr);
      }
    } else if (op.type === 'add') {
      if (!String(fields.addr || '').trim()) { errors.push('第 ' + (idx + 1) + ' 条：缺地址'); return; }
      row = findRow(fields.addr);
      if (!row) {
        row = rows.length ? Math.max.apply(null, rows.map(function (h) { return h.r; })) + 1 : 2;
        if (sh.getRange(row, map.addr).getDisplayValue()) row = sh.getLastRow() + 1;
      }
    } else {
      errors.push('第 ' + (idx + 1) + ' 条：未知操作 ' + op.type); return;
    }

    Object.keys(fields).forEach(function (f) {
      if (!map[f]) return;
      writeCell_(sh, row, map[f], String(fields[f]).trim());
    });
  });

  SpreadsheetApp.flush();
  var res = housesRead_();
  if (errors.length) { res.ok = false; res.error = errors.join('；'); }
  return res;
}

/** 在编辑器里手动运行一次，用来授权 Maps 服务、预热坐标缓存 */
function warmHouseGeo() {
  var r = housesRead_();
  Logger.log(r.houses.filter(function (h) { return h.lat != null; }).length + ' / ' + r.houses.length + ' 已定位');
  if (r.geoError) Logger.log('地理编码出错：' + r.geoError);
}
