/**
 * build-data.js — 本地数据生成脚本（完整版）
 * 读取 rawdata/ 中的 Excel，输出 data-inline.js（内置数据）
 * 
 * 支持：Target.xlsx 目标数据、AM维度、同比、达成率
 */

var XLSX = require('xlsx');
var fs = require('fs');
var path = require('path');

// ============================================================
// 配置
// ============================================================
var RAWDATA_DIR = 'D:\\VS Code\\fucaso-dashboard\\rawdata';

// YTD 月份控制：0=自动（当月），可手动指定 1-12
var YTD_MONTH = 0;
var rptM = parseInt(process.argv[2]) || YTD_MONTH || 0;
if (rptM > 12 || rptM < 1) rptM = 0;
if (!rptM) rptM = new Date().getMonth() + 1;

console.log('报表月份:', rptM);

// ============================================================
// 读取 Excel 文件
// ============================================================
var files = fs.readdirSync(RAWDATA_DIR);
var bsFile = files.find(f => f.startsWith('bs_order') && !f.startsWith('~$'));
if (!bsFile) { console.error('❌ 未找到 bs_order 文件'); process.exit(1); }

var bsWb = XLSX.readFile(path.join(RAWDATA_DIR, bsFile));
var bsRows = XLSX.utils.sheet_to_json(bsWb.Sheets[bsWb.SheetNames[0]], { header: 1 });

var mdWb = XLSX.readFile(path.join(RAWDATA_DIR, 'masterdata.xlsx'));
var mdRows = XLSX.utils.sheet_to_json(mdWb.Sheets[mdWb.SheetNames[0]]);

// 读取 Target.xlsx
var targetWb = XLSX.readFile(path.join(RAWDATA_DIR, 'Target.xlsx'));
var targetCompany = XLSX.utils.sheet_to_json(targetWb.Sheets['公司目标']);
var targetChallenge = XLSX.utils.sheet_to_json(targetWb.Sheets['挑战目标']);

// ============================================================
// 辅助函数
// ============================================================
function toLocal(d) {
  return d.getFullYear() + '-' +
    String(d.getMonth() + 1).padStart(2, '0') + '-' +
    String(d.getDate()).padStart(2, '0');
}

function excelToDate(v) {
  if (typeof v === 'number') return toLocal(new Date((v - 25569) * 86400 * 1000));
  if (typeof v === 'string') { var m = v.match(/(\d{4}-\d{2}-\d{2})/); return m ? m[1] : null; }
  return null;
}

function parseDt(v) {
  if (!v) return null;
  var m = String(v).match(/(\d{4}-\d{2}-\d{2})/);
  return m ? m[1] : null;
}

function inRange(d, f, t) { return d && d >= f && d <= t; }
function fmtNum(n) { return n.toLocaleString(); }

function maskName(name) {
  if (!name) return '';
  name = String(name).trim();
  if (name.length <= 1) return name + '*';
  if (/[a-zA-Z]/.test(name) && name.includes(' ')) {
    return name.split(' ').map(function(seg) {
      if (seg.length <= 2) return seg[0] + '*';
      return seg[0] + '*' + seg[seg.length - 1];
    }).join(' ');
  }
  if (/[a-zA-Z]/.test(name)) {
    if (name.length <= 2) return name[0] + '*';
    return name[0] + '*' + name[name.length - 1];
  }
  if (name.length === 2) return name[0] + '*';
  return name[0] + '*' + name[name.length - 1];
}

// ============================================================
// 主数据映射（扩展 AM/Area/Region）
// ============================================================
var masterMap = {};
mdRows.forEach(r => {
  var c = String(r['细胞追溯系统代码'] || '').trim();
  if (c) masterMap[c] = {
    name: String(r['标准医院名称'] || '').trim(),
    prov: String(r['省份'] || '').trim(),
    am: String(r['AM'] || '').trim(),
    area: String(r['Area'] || '').trim(),
    region: String(r['Region'] || '').trim()
  };
});

// ============================================================
// 解析 bs_order
// ============================================================
var headers = bsRows[1];
var ci = {};
headers.forEach((h, i) => {
  h = String(h || '').replace(/\n/g, '');
  if (h === '医疗机构编码') ci.org = i;
  if (h === '医疗机构名称') ci.orgName = i;
  if (h === '患者姓名') ci.patient = i;
  if (h.includes('合同创建') && h.includes('日期')) ci.od = i;
  if (h.includes('实际回输') && h.includes('结束时间')) ci.re = i;
  if (h.includes('实际单采') && h.includes('开始时间')) ci.ap = i;
  if (h.includes('生产质量') && h.includes('放行时间')) ci.qa = i;
});

var records = [];
var unmatchedWarnings = [];

for (var i = 2; i < bsRows.length; i++) {
  var row = bsRows[i]; if (!row) continue;
  var od = excelToDate(row[ci.od]), re = parseDt(row[ci.re]), ap = parseDt(row[ci.ap]), qa = parseDt(row[ci.qa]);
  var code1 = String(row[ci.org] || '').trim();
  var code2 = String(row[5] || '').trim();
  var m = null;

  if (code1 && masterMap[code1]) {
    m = masterMap[code1];
  } else if (code2 && masterMap[code2]) {
    m = masterMap[code2];
  }

  var hosp, prov, am, area, region;
  if (m) {
    hosp = m.name; prov = m.prov; am = m.am; area = m.area; region = m.region;
  } else {
    var fallbackHosp = String(row[ci.orgName] || row[6] || '').trim();
    if (code1 || code2) {
      unmatchedWarnings.push({ row: i, code1: code1, code2: code2, rawHosp: fallbackHosp || row[6] || '' });
    }
    hosp = fallbackHosp; prov = ''; am = ''; area = ''; region = '';
  }

  if (!hosp) continue;
  var patient = maskName(row[ci.patient]);
  var isSG = (prov === '新加坡' || prov.includes('新加坡'));
  records.push({ hosp, prov: isSG ? '' : prov, am, area, region, od, re, ap, qa, noMap: isSG, patient });
}

if (unmatchedWarnings.length > 0) {
  console.error('\n========================================');
  console.error('⚠️  警告：发现 ' + unmatchedWarnings.length + ' 条记录无法匹配主数据！');
  var seen = new Set();
  unmatchedWarnings.forEach(function(w) {
    var key = w.code1 || w.code2;
    if (!seen.has(key)) { seen.add(key); console.error('  · 编码: ' + key + '  原始医院名: ' + (w.rawHosp || '未知')); }
  });
  console.error('请将以上编码添加至 masterdata.xlsx 后重新运行！');
  console.error('========================================\n');
  process.exit(1);
}

// ============================================================
// 时间基准
// ============================================================
var DP = toLocal(new Date());
var Y = DP.slice(0, 4);
var dpM = DP.slice(5, 7);
var dpM0 = parseInt(dpM);

// YTD 范围：当年1月1日 到 rptM月末（或今天，取较早者）
var ytdEnd = Y + '-' + String(rptM).padStart(2, '0') + '-31';
if (ytdEnd > DP) ytdEnd = DP;

// MTD 范围：rptM月1日 到 rptM月末（或今天）
var mtdStart = Y + '-' + String(rptM).padStart(2, '0') + '-01';
var mtdEnd = Y + '-' + String(rptM).padStart(2, '0') + '-31';
if (mtdEnd > DP) mtdEnd = DP;

// ============================================================
// 基础统计
// ============================================================
var ytdO = records.filter(r => inRange(r.od, Y + '-01-01', ytdEnd)).length;
var ytdR = records.filter(r => inRange(r.re, Y + '-01-01', ytdEnd)).length;
var mtdO = records.filter(r => inRange(r.od, mtdStart, mtdEnd)).length;
var mtdR = records.filter(r => inRange(r.re, mtdStart, mtdEnd)).length;

var monO = [0,0,0,0,0,0,0,0,0,0,0,0], monR = [0,0,0,0,0,0,0,0,0,0,0,0];
records.forEach(r => {
  if (r.od && r.od >= Y + '-01-01' && r.od <= ytdEnd) { var mi = parseInt(r.od.slice(5,7)) - 1; if (mi < 12) monO[mi]++; }
  if (r.re && r.re >= Y + '-01-01' && r.re <= ytdEnd) { var mi = parseInt(r.re.slice(5,7)) - 1; if (mi < 12) monR[mi]++; }
});

// 过去7天
var dpDate = new Date(DP.replace(/-/g, '/'));
var p7 = [];
for (var d = 6; d >= 0; d--) {
  var dt = new Date(dpDate); dt.setDate(dt.getDate() - d);
  var ds = toLocal(dt);
  var oMap = {}, rMap = {}, aMap = {}, qMap = {};
  records.forEach(r => {
    if (r.od === ds) { var key = r.hosp + '|||' + r.patient; oMap[key] = (oMap[key]||0)+1; }
    if (r.re === ds) { var key = r.hosp + '|||' + r.patient; rMap[key] = (rMap[key]||0)+1; }
    if (r.ap === ds) { var key = r.hosp + '|||' + r.patient; aMap[key] = (aMap[key]||0)+1; }
    if (r.qa === ds) { var key = r.hosp + '|||' + r.patient; qMap[key] = (qMap[key]||0)+1; }
  });
  p7.push({ date: ds, orders: oMap, reinfusion: rMap, apheresis: aMap, quality: qMap });
}
p7.reverse();

var p7totalO = 0, p7totalR = 0, p7totalA = 0, p7totalQ = 0;
p7.forEach(d => {
  p7totalO += Object.values(d.orders).reduce((a,b)=>a+b,0);
  p7totalR += Object.values(d.reinfusion).reduce((a,b)=>a+b,0);
  p7totalA += Object.values(d.apheresis).reduce((a,b)=>a+b,0);
  p7totalQ += Object.values(d.quality).reduce((a,b)=>a+b,0);
});

// 省份数据
var provGeoMap = { '香港':'香港特别行政区','澳门':'澳门特别行政区','台湾':'台湾省','新加坡':'' };
var provData = {};
records.forEach(r => {
  if (r.noMap) return;
  if (r.od && r.od >= Y+'-01-01' && r.od <= ytdEnd) {
    var p = provGeoMap[r.prov] !== undefined ? provGeoMap[r.prov] : r.prov;
    if (p) provData[p] = (provData[p]||0)+1;
  }
});

var ALL_P = ['北京市','天津市','上海市','重庆市','河北省','山西省','辽宁省','吉林省','黑龙江省','江苏省','浙江省','安徽省','福建省','江西省','山东省','河南省','湖北省','湖南省','广东省','广西壮族自治区','海南省','四川省','贵州省','云南省','西藏自治区','陕西省','甘肃省','青海省','宁夏回族自治区','新疆维吾尔自治区','内蒙古自治区','台湾省','香港特别行政区','澳门特别行政区'];
var mapJson = ALL_P.map(p => ({ name: p, value: provData[p]||0 }));
var provRank5 = Object.entries(provData).sort((a,b) => b[1]-a[1]).slice(0, 5);

var topProvNames = provRank5.slice(0,3).map(e =>
  e[0].replace('北京市','北京').replace('天津市','天津').replace('上海市','上海')
     .replace('特别行政区','').replace('壮族自治区','').replace('省','').replace('市','')
);

var todayO = records.filter(r => r.od === DP).length;
var todayR = records.filter(r => r.re === DP).length;
var todayA = records.filter(r => r.ap === DP).length;
var todayQ = records.filter(r => r.qa === DP).length;

// ============================================================
// Target 目标数据解析
// ============================================================
// 公司目标：{YM: {下单, 回输}}
var companyTargetMap = {};
targetCompany.forEach(r => {
  var ym = String(r['YM'] || '');
  if (ym) companyTargetMap[ym] = { o: parseFloat(r['下单']) || 0, r: parseFloat(r['回输']) || 0 };
});

// 挑战目标：{AM|YM: {下单, 回输}}
var challengeTargetMap = {};
targetChallenge.forEach(r => {
  var ym = String(r['YM'] || '');
  var am = String(r['AM'] || '');
  if (ym && am) challengeTargetMap[am + '|' + ym] = { o: parseFloat(r['下单']) || 0, r: parseFloat(r['回输']) || 0 };
});

// 计算 YTD/MTD 目标
var kpiYtdTargetO = 0, kpiYtdTargetR = 0;
var kpiMtdTargetO = 0, kpiMtdTargetR = 0;
for (var m = 1; m <= rptM; m++) {
  var ym = Y + String(m).padStart(2, '0');
  if (companyTargetMap[ym]) {
    kpiYtdTargetO += companyTargetMap[ym].o;
    kpiYtdTargetR += companyTargetMap[ym].r;
  }
}
var mtdYM = Y + String(rptM).padStart(2, '0');
if (companyTargetMap[mtdYM]) {
  kpiMtdTargetO = companyTargetMap[mtdYM].o;
  kpiMtdTargetR = companyTargetMap[mtdYM].r;
}

// 达成率
var kpiYtdRateO = kpiYtdTargetO > 0 ? Math.round(ytdO / kpiYtdTargetO * 100) : 0;
var kpiYtdRateR = kpiYtdTargetR > 0 ? Math.round(ytdR / kpiYtdTargetR * 100) : 0;
var kpiMtdRateO = kpiMtdTargetO > 0 ? Math.round(mtdO / kpiMtdTargetO * 100) : 0;
var kpiMtdRateR = kpiMtdTargetR > 0 ? Math.round(mtdR / kpiMtdTargetR * 100) : 0;

// 同比：从records中筛选去年同期数据
var lyY = String(parseInt(Y) - 1);
var lyDP = lyY + DP.slice(4); // 去年同期日期
var lyMtdStart = lyY + '-' + String(rptM).padStart(2, '0') + '-01';
var lyYtdO = records.filter(r => inRange(r.od, lyY + '-01-01', lyDP)).length;
var lyYtdR = records.filter(r => inRange(r.re, lyY + '-01-01', lyDP)).length;
var lyMtdO = records.filter(r => inRange(r.od, lyMtdStart, lyDP)).length;
var lyMtdR = records.filter(r => inRange(r.re, lyMtdStart, lyDP)).length;

function calcYoy(curr, last) {
  if (last === null || last === undefined || last === 0) return null;
  return Math.round((curr - last) / last * 100);
}

var kpiYtdYoyO = calcYoy(ytdO, lyYtdO);
var kpiYtdYoyR = calcYoy(ytdR, lyYtdR);
var kpiMtdYoyO = calcYoy(mtdO, lyMtdO);
var kpiMtdYoyR = calcYoy(mtdR, lyMtdR);

// ============================================================
// AM 销售达成表数据
// ============================================================
// Region 顺序：DOM → HK → SG → KSA → 其他
// DOM 内 AM 固定顺序
var DOM_AM_ORDER = ['崔珺', '赵蕊', '赵俊兴', '龚卉', '高威龙', '董硕', '兰明金', '李磊'];
var REGION_ORDER = ['DOM', 'HK', 'SG', 'KSA'];

// 收集所有 AM 数据
var amSet = new Set();
records.forEach(r => { if (r.am) amSet.add(r.am); });

// 按 Region 分组 AM
var amByRegion = {};
records.forEach(r => {
  if (!r.am) return;
  var rg = r.region || '其他';
  if (!amByRegion[rg]) amByRegion[rg] = new Set();
  amByRegion[rg].add(r.am);
});

// 排序函数
function sortRegion(a, b) {
  var ia = REGION_ORDER.indexOf(a);
  var ib = REGION_ORDER.indexOf(b);
  if (ia >= 0 && ib >= 0) return ia - ib;
  if (ia >= 0) return -1;
  if (ib >= 0) return 1;
  return a.localeCompare(b);
}

function sortAM(amList, region) {
  if (region === 'DOM') {
    return amList.sort((a, b) => {
      var ia = DOM_AM_ORDER.indexOf(a);
      var ib = DOM_AM_ORDER.indexOf(b);
      if (ia >= 0 && ib >= 0) return ia - ib;
      if (ia >= 0) return -1;
      if (ib >= 0) return 1;
      return a.localeCompare(b);
    });
  }
  return amList.sort((a, b) => a.localeCompare(b));
}

// 构建 AM 表数据
function buildAmTable(type) {
  // type: 'od' 或 're'
  var dateField = type === 'od' ? 'od' : 're';
  var regions = Object.keys(amByRegion).sort(sortRegion);
  var rows = [];

  regions.forEach(rg => {
    var amList = sortAM(Array.from(amByRegion[rg]), rg);
    var rgHasSum = (rg === 'DOM' && amList.length > 1);
    var rgMonthly = {};
    var rgYtdActual = 0;
    var rgYtdChallenge = 0;

    amList.forEach(am => {
      // 每月实际值
      var monthly = new Array(12).fill(0);
      records.forEach(r => {
        if (r.am !== am || !r[dateField]) return;
        var dt = r[dateField];
        if (dt >= Y + '-01-01' && dt <= ytdEnd) {
          var mi = parseInt(dt.slice(5,7)) - 1;
          if (mi < 12) monthly[mi]++;
        }
      });

      // YTD 实际
      var ytdActual = monthly.reduce((a,b) => a+b, 0);

      // YTD 挑战目标
      var ytdChallenge = 0;
      for (var m = 1; m <= rptM; m++) {
        var key = am + '|' + Y + String(m).padStart(2, '0');
        if (challengeTargetMap[key]) ytdChallenge += challengeTargetMap[key][type === 'od' ? 'o' : 'r'];
      }

      // 达成率
      var rate = ytdChallenge > 0 ? Math.round(ytdActual / ytdChallenge * 100) : (ytdChallenge === 0 ? null : 0);

      // 同比（暂无去年数据）
      var yoy = null;

      rows.push({
        region: rg, am: am, monthly: monthly,
        ytdActual: ytdActual, ytdChallenge: ytdChallenge,
        rate: rate, yoy: yoy, isSum: false
      });

      // 汇总
      if (rgHasSum) {
        for (var i = 0; i < 12; i++) rgMonthly[i] = (rgMonthly[i] || 0) + monthly[i];
        rgYtdActual += ytdActual;
        rgYtdChallenge += ytdChallenge;
      }
    });

    // DOM 汇总行
    if (rgHasSum) {
      var rgRate = rgYtdChallenge > 0 ? Math.round(rgYtdActual / rgYtdChallenge * 100) : null;
      rows.push({
        region: rg, am: rg + '汇总', monthly: Object.values(rgMonthly),
        ytdActual: rgYtdActual, ytdChallenge: rgYtdChallenge,
        rate: rgRate, yoy: null, isSum: true
      });
    }
  });

  return rows;
}

var amOrderTable = buildAmTable('od');
var amReinfusionTable = buildAmTable('re');

// ============================================================
// 组装输出
// ============================================================
var summary = {
  DP, Y, dpM, dpM0, rptM,
  ytdO, ytdR, mtdO, mtdR,
  todayO, todayR, todayA, todayQ,
  p7totalO, p7totalR, p7totalA, p7totalQ,
  topProvNames,
  monO, monR,
  mapJson, provRank5,
  p7,

  // KPI with targets
  kpiYtdO: ytdO, kpiYtdR: ytdR,
  kpiYtdTargetO, kpiYtdTargetR,
  kpiYtdRateO, kpiYtdRateR,
  kpiYtdYoyO, kpiYtdYoyR,

  kpiMtdO: mtdO, kpiMtdR: mtdR,
  kpiMtdTargetO, kpiMtdTargetR,
  kpiMtdRateO, kpiMtdRateR,
  kpiMtdYoyO, kpiMtdYoyR,

  // AM tables
  amOrderTable, amReinfusionTable,

  // Header summary text (for display)
  headerSummary: 'YTD下单 ' + fmtNum(ytdO) + ' 单，回输 ' + fmtNum(ytdR) + ' 单；' + rptM + '月MTD 下单 ' + fmtNum(mtdO) + ' 单，回输 ' + fmtNum(mtdR) + ' 单'
};

var outPath = path.join(__dirname, 'data-inline.js');
var outContent = '/** 内置数据 — 由 build-data.js 自动生成 */\nwindow.BUILTIN_DATA = ' + JSON.stringify(summary, null, 2) + ';\n';
fs.writeFileSync(outPath, outContent, 'utf-8');

console.log('\n✅ 内置数据已生成:', outPath);
console.log('YTD下单:', ytdO, ' YTD回输:', ytdR);
console.log('YTD目标下单:', kpiYtdTargetO, ' YTD目标回输:', kpiYtdTargetR);
console.log('YTD达成率下单:', kpiYtdRateO + '%', ' YTD达成率回输:', kpiYtdRateR + '%');
console.log('MTD下单:', mtdO, ' MTD回输:', mtdR);
console.log('AM数量:', amOrderTable.length);
console.log('\n下一步：运行 node deploy.js 推送到 GitHub');
