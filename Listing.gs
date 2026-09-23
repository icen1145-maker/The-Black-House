/**
 * 看房板 · 去 Redfin 抓当前行情，直接在 Google 的服务器上跑
 *
 * 不依赖浏览器、不依赖 Claude、不依赖电脑开着。
 *   - 网页上每处房源的「从房源网站更新一次」按钮 → Houses.gs 的 refresh op → refreshRow_()
 *   - 每天自动跑一次 → dailyRefresh()（用 installDailyRefresh() 装触发器）
 *
 * 只写 2026 tab 的 价格 / 售价 / 状态 / Open House 四列，外加「房源属性」tab 里还空着的硬数据。
 */

var UA = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36';
var REDFIN = 'https://www.redfin.com';
var FIELD_CN = { price: '挂牌价', sold: '成交价', status: '状态', oh: 'open house' };

function get_(url) {
  try {
    var r = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true, followRedirects: true,
      headers: { 'User-Agent': UA, 'Accept': 'text/html,application/xhtml+xml', 'Accept-Language': 'en-US,en;q=0.9' }
    });
    return r.getResponseCode() === 200 ? r.getContentText() : '';
  } catch (e) { return ''; }
}

/* ---------- 地址 → 房源页 URL（找到就记在脚本属性里，下次不用再找） ---------- */
function listingUrl_(addr, force) {
  var props = PropertiesService.getScriptProperties();
  var k = 'url:' + normAddr_(addr).slice(0, 180);
  if (!force) { var hit = props.getProperty(k); if (hit) return hit; }
  var url = findListingUrl_(addr);
  if (url) props.setProperty(k, url);
  return url;
}

function findListingUrl_(addr) {
  var a = String(addr || '').replace(/\s+/g, ' ').trim();
  var m = a.match(/^(\d+[A-Za-z]?)\s+(.+?),\s*[^,]+,\s*CA\s*(\d{5})/i);
  if (!m) return '';
  var num = m[1], street = m[2].trim(), zip = m[3];
  var want = normAddr_(num + ' ' + street);
  // sitemap 上方位字母（N/S/E/W）常被省掉，两种都试
  var tries = [street];
  var bare = street.replace(/^([NSEW])\s+/i, '');
  if (bare !== street) tries.push(bare);

  for (var i = 0; i < tries.length; i++) {
    var html = get_(REDFIN + '/sitemap/CA/' + zip + '/street/' + encodeURIComponent(tries[i]).replace(/%20/g, '+'));
    if (!html) continue;
    var re = /<a[^>]+href="(\/[A-Z]{2}\/[^"]+\/home\/\d+)"[^>]*>([^<]{4,80})<\/a>/g, x;
    while ((x = re.exec(html))) {
      if (normAddr_(x[2]).indexOf(want) === 0) return REDFIN + x[1];
    }
  }
  return '';
}

/* ---------- 解析房源页 ---------- */
function fetchListing_(addr, force) {
  var url = listingUrl_(addr, force);
  if (!url) return { error: '在 Redfin 上没找到这个地址' };
  var html = get_(url);
  if (!html) {
    if (!force) return fetchListing_(addr, true);   // 记的链接可能过期了，重新找一次
    return { error: 'Redfin 没返回页面', url: url };
  }
  var out = { url: url };

  var desc = (html.match(/<meta name="description" content="([^"]{0,600})/) || [])[1] || '';
  var band = (html.match(/StatusBannerSection--statusDot[^>]*><\/div>([^<]{1,40})</) || [])[1] || '';
  var mls = (html.match(/mlsStatusDisplay[^{]{0,12}\{[^"]{0,12}"?displayValue\\?":\\?"([^"\\]{1,40})/) || [])[1] || '';
  var raw = band || mls || desc.split(':')[0];
  out.raw = raw;

  var s = String(raw).toLowerCase();
  if (/sold|closed/.test(s)) out.status = 'Sold';
  else if (/pending|contingent|backup|under contract/.test(s)) out.status = 'Pending';
  else if (/for sale|active|coming soon|new listing/.test(s)) out.status = 'On Sale';

  var num = function (t) { return t ? Number(String(t).replace(/,/g, '')) : null; };

  if (out.status === 'Sold') {
    out.sold = num((desc.match(/sold for \$([\d,]+)/i) || [])[1]);
    out.soldDate = (desc.match(/ on ([A-Z][a-z]{2} \d{1,2}, \d{4})/) || [])[1] || '';
  } else {
    out.price = num((html.match(/"offers":\{[^}]*"price":(\d+)/) || [])[1]) ||
                num((desc.match(/\$([\d,]+)/) || [])[1]);
  }

  // 顺手把硬数据也读出来（只在表格里还空着时才写）
  var f = desc.match(/(\d+)\s*beds?[,\s∙]+([\d.]+)\s*baths?[,\s∙]+([\d,]+)\s*sq\.?\s*ft/i);
  if (f) { out.bd = num(f[1]); out.ba = Number(f[2]); out.sqft = num(f[3]); }

  return out;
}

/* ---------- 价格格式：整千写简写，否则写完整数字 ---------- */
function money_(n) {
  n = Number(n);
  if (!n || !isFinite(n)) return '';
  if (n % 1000 !== 0) return String(Math.round(n));
  if (n >= 1e6) return String(n / 1e6).replace(/(\.\d*[1-9])0+$/, '$1').replace(/\.0+$/, '') + 'm';
  return (n / 1000) + 'k';
}

/* 比较时按数值比，"1.105m" 和 "1105000" 算一样，避免反复重写 */
function sameVal_(a, b) {
  var norm = function (t) {
    var s = String(t == null ? '' : t).toLowerCase().replace(/[$,\s]/g, '');
    var m = s.match(/^([\d.]+)(m|k)?$/);
    if (!m) return s;
    var n = parseFloat(m[1]);
    if (m[2] === 'm') n *= 1e6;
    else if (m[2] === 'k') n *= 1e3;
    else if (n < 100) n *= 1e6;
    else if (n < 10000) n *= 1e3;
    return String(Math.round(n));
  };
  return norm(a) === norm(b);
}

/* ---------- 刷新一行 ---------- */
function refreshRow_(sh, map, h) {
  var info = fetchListing_(h.addr);
  if (info.error) return { addr: h.addr, note: info.error, err: true };
  var changed = [];
  var put = function (field, val) {
    if (val == null || val === '') return;
    if (!map[field]) return;
    if (sameVal_(h[field], val)) return;
    writeCell_(sh, h.r, map[field], String(val));
    changed.push(FIELD_CN[field] + ' ' + (String(h[field] || '').trim() || '空') + ' → ' + val);
  };

  if (info.status) put('status', info.status);
  if (info.status === 'Sold') {
    // 成交了只补成交价；挂牌价保持之前记的那个（就是成交前最后的挂牌价）
    if (info.sold) put('sold', money_(info.sold));
  } else if (info.price) {
    put('price', money_(info.price));
  }
  // 不在卖了就把 open house 清掉
  if (info.status && info.status !== 'On Sale' && String(h.oh || '').trim() && map.oh) {
    writeCell_(sh, h.r, map.oh, '');
    changed.push('清掉 open house');
  }
  // 硬数据只补空的，不覆盖手填
  var had = attrAll_()[normAddr_(h.addr)] || {};
  var facts = {};
  ['bd', 'ba', 'sqft'].forEach(function (k) { if (info[k] && !String(had[k] || '').trim()) facts[k] = info[k]; });
  if (Object.keys(facts).length) { attrSave_(h.addr, facts); changed.push('补了卧室/卫生间/面积'); }

  return { addr: h.addr, url: info.url, raw: info.raw, note: changed.join('；'), changed: changed.length > 0 };
}

/* ---------- 每天自动跑 ---------- */
function dailyRefresh() {
  var sh = houseSheet_();
  var map = houseColMap_(sh);
  var rows = houseRows_(sh, map);
  var lines = [];
  rows.forEach(function (h) {
    var st = String(h.status || '').toLowerCase();
    // 已经成交而且成交价也记了的，不用再查
    if (/sold|closed/.test(st) && String(h.sold || '').trim()) return;
    var r = refreshRow_(sh, map, h);
    if (r.changed || r.err) lines.push(r.addr + '：' + r.note);
    Utilities.sleep(700);
  });
  SpreadsheetApp.flush();
  var msg = lines.length ? lines.join('\n') : '今天没有变化';
  PropertiesService.getScriptProperties().setProperty('lastDaily',
    Utilities.formatDate(new Date(), 'America/Los_Angeles', 'yyyy-MM-dd HH:mm') + '\n' + msg);
  Logger.log(msg);
}

/** 装 / 重装每天早上 9 点的触发器（在编辑器里手动跑一次即可） */
function installDailyRefresh() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'dailyRefresh') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyRefresh').timeBased().atHour(9).everyDays(1).create();
  Logger.log('已装好：每天早上 9 点自动更新');
}

/** 看上一次自动更新做了什么 */
function lastDailyLog() {
  Logger.log(PropertiesService.getScriptProperties().getProperty('lastDaily') || '还没跑过');
}
