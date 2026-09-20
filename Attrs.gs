/**
 * 看房板 · 房源属性（自动抓取 + 手填）
 *
 * 写在表格的「房源属性」tab 里，2026 tab 一列不动。
 *   A 地址 | B 卧室 | C 卫生间 | D 室内 sqft | E 地皮 sqft | F 建造年份 | G 电线
 *   H Superfund | I 噪音 | J 散步 / 公园 | K 风水提示 | L 朝向 | M 更新时间 | N 数据来源
 *
 * B–G 是「房子硬数据」，由网页写（粘贴房源描述自动解析，或手填），脚本不覆盖。
 * H–N 是「环境信息」，脚本自动算：OpenStreetMap（Overpass）+ EPA Superfund + 几何计算。
 * 新房源由定时触发器自动补齐：在编辑器里运行一次 setupAttrTrigger 即可。
 */

var ATTR_SHEET = '房源属性';
var ATTR_COLS = ['地址', '卧室', '卫生间', '室内 sqft', '地皮 sqft', '建造年份', '电线',
                 'Superfund', '噪音', '散步 / 公园', '风水提示', '朝向', '更新时间', '数据来源'];
var ATTR_KEYS = ['addr', 'bd', 'ba', 'sqft', 'lot', 'year', 'elec',
                 'sf', 'noise', 'walk', 'fs', 'face', 'at', 'src'];
var FACT_KEYS = ['bd', 'ba', 'sqft', 'lot', 'year', 'elec'];       // 网页写的
var ENV_KEYS  = ['sf', 'noise', 'walk', 'fs', 'face', 'at', 'src']; // 脚本写的
var ENRICH_MAX_AGE_DAYS = 180;
var ENRICH_PER_RUN = 8;

function attrSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ATTR_SHEET);
  if (!sh) {
    sh = ss.insertSheet(ATTR_SHEET);
    sh.getRange(1, 1, 1, ATTR_COLS.length).setValues([ATTR_COLS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 260);
    [8, 9, 10, 11].forEach(function (c) { sh.setColumnWidth(c, 220); });
  }
  return sh;
}

/** 整张表读成 {normAddr: {行号, 各字段}} */
function attrAll_() {
  var sh = attrSheet_();
  var last = sh.getLastRow();
  var out = {};
  if (last < 2) return out;
  var vals = sh.getRange(2, 1, last - 1, ATTR_COLS.length).getDisplayValues();
  for (var i = 0; i < vals.length; i++) {
    var a = String(vals[i][0] || '').trim();
    if (!a) continue;
    var o = { r: i + 2 };
    for (var c = 1; c < ATTR_KEYS.length; c++) o[ATTR_KEYS[c]] = String(vals[i][c] || '').trim();
    out[normAddr_(a)] = o;
  }
  return out;
}

/** 合并写入一行（只写传进来的字段，其余保持原样） */
function attrSave_(addr, obj) {
  var sh = attrSheet_();
  var all = attrAll_();
  var hit = all[normAddr_(addr)];
  var row = hit ? hit.r : Math.max(sh.getLastRow() + 1, 2);
  if (!hit) sh.getRange(row, 1).setNumberFormat('@').setValue(addr);
  for (var c = 1; c < ATTR_KEYS.length; c++) {
    var k = ATTR_KEYS[c];
    if (obj[k] === undefined) continue;
    sh.getRange(row, c + 1).setNumberFormat('@').setValue(String(obj[k] == null ? '' : obj[k]));
  }
}

/* ================= 抓取用的小工具 ================= */
function miles_(a, b, c, d) {
  var R = 3958.8, r = Math.PI / 180;
  var dLat = (c - a) * r, dLng = (d - b) * r;
  var h = Math.sin(dLat / 2) * Math.sin(dLat / 2) + Math.cos(a * r) * Math.cos(c * r) * Math.sin(dLng / 2) * Math.sin(dLng / 2);
  return 2 * R * Math.asin(Math.sqrt(h));
}
function meters_(a, b, c, d) { return miles_(a, b, c, d) * 1609.34; }
function bearing_(a, b, c, d) {
  var r = Math.PI / 180;
  var y = Math.sin((d - b) * r) * Math.cos(c * r);
  var x = Math.cos(a * r) * Math.sin(c * r) - Math.sin(a * r) * Math.cos(c * r) * Math.cos((d - b) * r);
  return (Math.atan2(y, x) / r + 360) % 360;
}
function compass_(deg) {
  var n = ['正北', '东北', '正东', '东南', '正南', '西南', '正西', '西北'];
  return n[Math.round(deg / 45) % 8];
}
function fmtM_(m) { return m < 1000 ? Math.round(m / 10) * 10 + ' 米' : (m / 1609.34).toFixed(1) + ' 英里'; }

function overpass_(q) {
  var hosts = ['https://overpass-api.de/api/interpreter',
               'https://overpass.kumi.systems/api/interpreter'];
  for (var i = 0; i < hosts.length; i++) {
    try {
      var res = UrlFetchApp.fetch(hosts[i], {
        method: 'post', payload: { data: q }, muteHttpExceptions: true,
        followRedirects: true, validateHttpsCertificates: true
      });
      if (res.getResponseCode() === 200) {
        var j = JSON.parse(res.getContentText());
        if (j && j.elements) return j.elements;
      }
    } catch (e) {}
    Utilities.sleep(800);
  }
  return null;
}

/** EPA Superfund（SEMS）点位，几个备选接口挨个试，哪个通用哪个 */
function superfundNear_(lat, lng, km) {
  var d = km / 111;
  var box = [lng - d, lat - d, lng + d, lat + d].join(',');
  var urls = [
    'https://geopub.epa.gov/arcgis/rest/services/EMEF/efpoints/MapServer/0/query?f=json&returnGeometry=true&outSR=4326&inSR=4326&spatialRel=esriSpatialRelIntersects&geometryType=esriGeometryEnvelope&outFields=*&geometry=' + encodeURIComponent(box),
    'https://services.arcgis.com/cJ9YHowT8TU7DUyn/arcgis/rest/services/Superfund_National_Priorities_List_NPL_Sites/FeatureServer/0/query?f=json&returnGeometry=true&outSR=4326&inSR=4326&spatialRel=esriSpatialRelIntersects&geometryType=esriGeometryEnvelope&outFields=*&geometry=' + encodeURIComponent(box)
  ];
  for (var i = 0; i < urls.length; i++) {
    try {
      var res = UrlFetchApp.fetch(urls[i], { muteHttpExceptions: true });
      if (res.getResponseCode() !== 200) continue;
      var j = JSON.parse(res.getContentText());
      if (!j || !j.features) continue;
      var out = [];
      j.features.forEach(function (f) {
        var g = f.geometry || {}, at = f.attributes || {};
        if (g.x == null || g.y == null) return;
        var name = at.SITE_NAME || at.NAME || at.PRIMARY_NAME || at.FAC_NAME || at.SITE_NM || '未命名场地';
        var npl = String(at.NPL_STATUS_CODE || at.NPL_STATUS || at.STATUS || '');
        out.push({ name: String(name), npl: npl, d: meters_(lat, lng, g.y, g.x) });
      });
      out.sort(function (a, b) { return a.d - b.d; });
      return { list: out, src: i === 0 ? 'EPA EMEF' : 'EPA NPL' };
    } catch (e) {}
  }
  return null;
}

/* ================= 核心：算一处房源的环境信息 ================= */
function enrichOne_(addr, lat, lng) {
  var q = '[out:json][timeout:30];(' +
    'way(around:1200,LAT,LNG)[leisure=park];relation(around:1200,LAT,LNG)[leisure=park];' +
    'way(around:1000,LAT,LNG)[landuse=cemetery];' +
    'node(around:700,LAT,LNG)[amenity~"^(fuel|hospital|place_of_worship)$"];' +
    'node(around:700,LAT,LNG)[shop=funeral_directors];' +
    'way(around:600,LAT,LNG)[power];node(around:250,LAT,LNG)[power~"^(pole|tower)$"];' +
    'way(around:1500,LAT,LNG)[railway~"^(rail|light_rail)$"];' +
    'way(around:1500,LAT,LNG)[highway~"^(motorway|trunk|primary|secondary)$"];' +
    'way(around:900,LAT,LNG)[highway~"^(footway|path)$"];' +
    'node(around:1200,LAT,LNG)[shop=supermarket];' +
    'way(around:1200,LAT,LNG)[natural=water];way(around:1200,LAT,LNG)[waterway=river];' +
    ');out center tags 300;' +
    'way(around:160,LAT,LNG)[highway][highway!~"^(footway|path|service|cycleway)$"];out geom 40;';
  q = q.replace(/LAT/g, lat).replace(/LNG/g, lng);

  var els = overpass_(q);
  var out = { at: new Date().toISOString().slice(0, 16).replace('T', ' ') };
  var srcs = [];

  if (els) {
    srcs.push('OSM');
    var near = [];     // {kind, name, d, tags}
    var roads = [];    // 带几何的近处马路
    els.forEach(function (e) {
      var t = e.tags || {};
      if (e.type === 'way' && e.geometry && e.geometry.length && t.highway) {
        roads.push({ t: t, g: e.geometry });
      }
      var c = e.center || (e.lat != null ? { lat: e.lat, lon: e.lon } : null);
      if (!c) return;
      near.push({ t: t, d: meters_(lat, lng, c.lat, c.lon), name: t.name || '' });
    });
    var pick = function (fn) {
      var hit = null;
      near.forEach(function (o) { if (fn(o.t) && (!hit || o.d < hit.d)) hit = o; });
      return hit;
    };
    var count = function (fn, within) {
      var n = 0;
      near.forEach(function (o) { if (fn(o.t) && o.d <= within) n++; });
      return n;
    };

    /* --- 散步 / 公园 --- */
    var park = pick(function (t) { return t.leisure === 'park'; });
    var parks800 = count(function (t) { return t.leisure === 'park'; }, 800);
    var trail = pick(function (t) { return t.highway === 'footway' || t.highway === 'path'; });
    var mkt = pick(function (t) { return t.shop === 'supermarket'; });
    var walk = [];
    if (park) walk.push('最近公园 ' + (park.name || '无名') + ' ' + fmtM_(park.d));
    else walk.push('1.2 公里内没有公园');
    if (parks800 > 1) walk.push('800 米内 ' + parks800 + " 个公园");
    if (trail && trail.d < 500) walk.push('有步道（' + fmtM_(trail.d) + '）');
    if (mkt) walk.push('超市 ' + fmtM_(mkt.d));
    out.walk = walk.join('；');

    /* --- 噪音 --- */
    var fw = pick(function (t) { return t.highway === 'motorway' || t.highway === 'trunk'; });
    var art = pick(function (t) { return t.highway === 'primary' || t.highway === 'secondary'; });
    var rail = pick(function (t) { return t.railway === 'rail' || t.railway === 'light_rail'; });
    var sjc = meters_(lat, lng, 37.3639, -121.9289);
    var score = 0, why = [];
    if (fw) {
      if (fw.d < 200) { score += 3; why.push('高速 ' + (fw.t.ref || '') + ' 就在 ' + fmtM_(fw.d)); }
      else if (fw.d < 500) { score += 2; why.push('离高速 ' + (fw.t.ref || '') + ' ' + fmtM_(fw.d)); }
      else if (fw.d < 1000) { score += 1; why.push('高速 ' + (fw.t.ref || '') + ' ' + fmtM_(fw.d)); }
    }
    if (art && art.d < 150) { score += 2; why.push('紧邻大路 ' + (art.name || art.t.ref || '')); }
    else if (art && art.d < 400) { score += 1; why.push('大路 ' + (art.name || art.t.ref || '') + ' ' + fmtM_(art.d)); }
    if (rail && rail.d < 400) { score += 2; why.push('铁路 ' + fmtM_(rail.d)); }
    else if (rail && rail.d < 800) { score += 1; why.push('铁路 ' + fmtM_(rail.d)); }
    // SJC 跑道大致 120°/300°，南北两端是进近走廊
    var brg = bearing_(37.3639, -121.9289, lat, lng);
    var inPath = (Math.abs(((brg - 300 + 540) % 360) - 180) < 25) || (Math.abs(((brg - 120 + 540) % 360) - 180) < 25);
    if (sjc < 5000 && inPath) { score += 2; why.push('在 SJC 进近方向 ' + fmtM_(sjc)); }
    else if (sjc < 3000) { score += 1; why.push('离 SJC ' + fmtM_(sjc)); }
    var label = score >= 5 ? '吵' : score >= 3 ? '偏吵' : score >= 1 ? '一般' : '安静';
    out.noise = label + (why.length ? '（' + why.join('；') + '）' : '（附近没有高速、大路、铁路）');

    /* --- 电线：OSM 上的电线杆/电线 --- */
    var pole = pick(function (t) { return t.power === 'pole' || t.power === 'tower'; });
    var line = pick(function (t) { return t.power === 'line' || t.power === 'minor_line'; });
    var sub = pick(function (t) { return t.power === 'substation'; });
    var eh = [];
    if (pole && pole.d < 120) eh.push('OSM 有电线杆（' + fmtM_(pole.d) + '）→ 多半是架空线');
    if (line && line.d < 300) eh.push((line.t.power === 'line' ? '高压线' : '配电线') + ' ' + fmtM_(line.d));
    if (sub && sub.d < 800) eh.push('变电站 ' + fmtM_(sub.d));
    out.elecHint = eh.length ? eh.join('；') : 'OSM 附近没有电线记录（不代表没有，最好看一眼街景）';

    /* --- 风水提示（机械判断，仅供参考） --- */
    var fs = [];
    var cem = pick(function (t) { return t.landuse === 'cemetery'; });
    var hosp = pick(function (t) { return t.amenity === 'hospital'; });
    var fune = pick(function (t) { return t.shop === 'funeral_directors'; });
    var fuel = pick(function (t) { return t.amenity === 'fuel'; });
    var wor = pick(function (t) { return t.amenity === 'place_of_worship'; });
    var water = pick(function (t) { return t.natural === 'water' || t.waterway === 'river'; });
    if (cem && cem.d < 800) fs.push('墓地 ' + fmtM_(cem.d));
    if (hosp && hosp.d < 400) fs.push('医院 ' + fmtM_(hosp.d));
    if (fune && fune.d < 600) fs.push('殡葬 ' + fmtM_(fune.d));
    if (fuel && fuel.d < 200) fs.push('加油站 ' + fmtM_(fuel.d));
    if (wor && wor.d < 200) fs.push('宗教场所 ' + fmtM_(wor.d));
    if (line && line.d < 150) fs.push('电线从头顶过');
    if (sub && sub.d < 300) fs.push('变电站很近');
    if (fw && fw.d < 200) fs.push('正贴高速');

    /* 朝向 + 路冲 + 死胡同：用近处马路的几何算 */
    var bestSeg = null;
    roads.forEach(function (rd) {
      for (var i = 0; i + 1 < rd.g.length; i++) {
        var p = rd.g[i], q2 = rd.g[i + 1];
        var mid = { lat: (p.lat + q2.lat) / 2, lon: (p.lon + q2.lon) / 2 };
        var d = meters_(lat, lng, mid.lat, mid.lon);
        if (!bestSeg || d < bestSeg.d) bestSeg = { d: d, p: p, q: q2, t: rd.t, g: rd.g };
      }
    });
    if (bestSeg && bestSeg.d < 120) {
      var mid2 = { lat: (bestSeg.p.lat + bestSeg.q.lat) / 2, lon: (bestSeg.p.lon + bestSeg.q.lon) / 2 };
      var face = bearing_(lat, lng, mid2.lat, mid2.lon);
      out.face = compass_(face) + '（门朝马路，' + Math.round(face) + '°）';
      if (face > 112 && face < 248) fs.push('朝南向阳');
      else if (face > 292 || face < 68) fs.push('朝北，冬天少晒太阳');
    }
    // 路冲：有马路的端点冲着房子
    roads.forEach(function (rd) {
      if (rd.g.length < 2) return;
      [[rd.g[rd.g.length - 1], rd.g[rd.g.length - 2]], [rd.g[0], rd.g[1]]].forEach(function (pair) {
        var endP = pair[0], prev = pair[1];
        var d = meters_(lat, lng, endP.lat, endP.lon);
        if (d > 90 || d < 3) return;
        var along = bearing_(prev.lat, prev.lon, endP.lat, endP.lon);
        var toHouse = bearing_(endP.lat, endP.lon, lat, lng);
        var diff = Math.abs(((along - toHouse + 540) % 360) - 180);
        if (diff < 22) fs.push('路口正对房子（路冲，' + fmtM_(d) + '）');
      });
    });
    if (water && water.d < 400) fs.push('近水（' + fmtM_(water.d) + '，同时留意水患）');
    out.fs = fs.length ? fs.join('；') : '附近没查到常说的忌讳（墓地/医院/加油站/高压线/路冲）';
  } else {
    out.noise = out.walk = out.fs = '（OSM 没抓到，稍后自动重试）';
  }

  /* --- Superfund --- */
  var sf = superfundNear_(lat, lng, 5);
  if (sf) {
    srcs.push(sf.src);
    var l = sf.list.filter(function (o) { return o.d < 5000; });
    if (!l.length) out.sf = '5 公里内没有 EPA 记录在案的场地';
    else {
      var top = l.slice(0, 3).map(function (o) { return o.name + ' ' + fmtM_(o.d) + (o.npl ? '（' + o.npl + '）' : ''); });
      var closest = l[0].d;
      out.sf = (closest < 1000 ? '⚠︎ ' : '') + l.length + ' 处：' + top.join('；');
    }
  } else {
    out.sf = '（EPA 没查到，稍后自动重试）';
  }

  out.src = srcs.join(' + ') || '—';
  return out;
}

/* ================= 批量 / 定时 ================= */
function attrStale_(a) {
  if (!a || !a.at) return true;
  var t = Date.parse(String(a.at).replace(' ', 'T') + ':00Z');
  if (isNaN(t)) return true;
  return (Date.now() - t) / 864e5 > ENRICH_MAX_AGE_DAYS;
}

/** 给还没算过环境信息的房源补齐，每次最多 ENRICH_PER_RUN 处 */
function enrichPending_(limit) {
  limit = limit || ENRICH_PER_RUN;
  var data = housesRead_();
  var all = attrAll_();
  var t0 = Date.now(), done = 0;
  for (var i = 0; i < data.houses.length && done < limit; i++) {
    var h = data.houses[i];
    if (h.lat == null) continue;
    var a = all[normAddr_(h.addr)];
    if (!attrStale_(a)) continue;
    var env = enrichOne_(h.addr, h.lat, h.lng);
    if (env.elecHint) {                       // 电线只在用户没填过的时候写提示
      if (!a || !a.elec) env.elec = env.elecHint;
      delete env.elecHint;
    }
    attrSave_(h.addr, env);
    done++;
    if (Date.now() - t0 > 240000) break;      // 4 分钟就收手，留余量
  }
  SpreadsheetApp.flush();
  return done;
}

/** 手动跑一次：把所有缺的补上（一次最多 8 处，多跑几次） */
function enrichHousesNow() {
  var n = enrichPending_(ENRICH_PER_RUN);
  Logger.log('这次补了 ' + n + ' 处；还缺的下次再跑或等定时触发器');
}

/** 建定时触发器：每 6 小时自动给新房源补属性。在编辑器里运行一次就行 */
function setupAttrTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'enrichHousesAuto') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('enrichHousesAuto').timeBased().everyHours(6).create();
  Logger.log('定时触发器已建好：每 6 小时补一次新房源属性');
}
function enrichHousesAuto() { enrichPending_(ENRICH_PER_RUN); }
