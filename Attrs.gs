/**
 * 看房板 · 房源属性表
 *
 * 只负责一件事：把属性读写到表格的「房源属性」tab。2026 tab 一列不动。
 *   A 地址 | B 卧室 | C 卫生间 | D 室内 sqft | E 地皮 sqft | F 建造年份 | G 电线
 *   H Superfund | I 噪音 | J 散步 / 公园 | K 风水提示 | L 朝向 | M 更新时间 | N 数据来源
 *
 * B–G「房子硬数据」：网页写（粘贴房源描述自动解析，或手填）。
 * H–N「环境信息」：网页端自己去 OpenStreetMap（Overpass）和 EPA Superfund 查，
 *   算完把结果发回来存这儿。放在浏览器里跑，脚本就不需要「访问外部网址」的授权。
 */

var ATTR_SHEET = '房源属性';
var ATTR_COLS = ['地址', '卧室', '卫生间', '室内 sqft', '地皮 sqft', '建造年份', '电线',
                 'Superfund', '噪音', '散步 / 公园', '风水提示', '朝向', '更新时间', '数据来源'];
var ATTR_KEYS = ['addr', 'bd', 'ba', 'sqft', 'lot', 'year', 'elec',
                 'sf', 'noise', 'walk', 'fs', 'face', 'at', 'src'];
var FACT_KEYS = ['bd', 'ba', 'sqft', 'lot', 'year', 'elec'];        // 网页手填 / 粘贴解析
var ENV_KEYS  = ['sf', 'noise', 'walk', 'fs', 'face', 'elec', 'at', 'src'];  // 网页抓取

function attrSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ATTR_SHEET);
  if (!sh) {
    sh = ss.insertSheet(ATTR_SHEET);
    sh.getRange(1, 1, 1, ATTR_COLS.length).setValues([ATTR_COLS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 260);
    [8, 9, 10, 11].forEach(function (c) { sh.setColumnWidth(c, 230); });
  }
  return sh;
}

/** 整张表读成 {normAddr: {r:行号, ...字段}} */
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

/** 合并写一行：只写传进来的字段，其余保持原样 */
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

/** 改地址时属性行跟着搬 */
function attrRename_(oldAddr, newAddr) {
  var all = attrAll_();
  var hit = all[normAddr_(oldAddr)];
  if (!hit) return;
  attrSheet_().getRange(hit.r, 1).setNumberFormat('@').setValue(newAddr);
}
