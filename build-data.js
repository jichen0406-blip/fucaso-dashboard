/**
 * build-data.js — 本地数据生成脚本
 * 读取 rawdata/ 中的 Excel，输出 data-inline.js（内置数据）
 * 
 * 用法：node build-data.js
 * 输出：data-inline.js（供网页直接引用）
 */

var XLSX = require('xlsx');
var fs = require('fs');
var path = require('path');

// ============================================================
// 数据处理（与 dashboard.js 保持逻辑一致）
// ============================================================
var rawDir = 'D:\\VS Code\\fucaso-dashboard\\rawdata';
var files = fs.readdirSync(rawDir);
var bsFile = files.find(f => f.startsWith('bs_order') && !f.startsWith('~$'));
if (!bsFile) { console.error('❌ 未找到 bs_order 文件'); process.exit(1); }

var bsWb = XLSX.readFile(path.join(rawDir, bsFile));
var bsRows = XLSX.utils.sheet_to_json(bsWb.Sheets[bsWb.SheetNames[0]], { header: 1 });
var mdWb = XLSX.readFile(path.join(rawDir, 'masterdata.xlsx'));
var mdRows = XLSX.utils.sheet_to_json(mdWb.Sheets[mdWb.SheetNames[0]]);

var masterMap = {};
mdRows.forEach(r => {
  var c = String(r['细胞追溯系统代码'] || '').trim();
  if (c) masterMap[c] = {
    name: String(r['标准医院名称'] || '').trim(),
    prov: String(r['省份'] || '').trim()
  };
});

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

var headers = bsRows[1];
var ci = {};
headers.forEach((h, i) => {
  h = String(h || '').replace(/\n/g, '');
  if (h === '医疗机构编码') ci.org = i;
  if (h === '医疗机构名称') ci.orgName = i;
  if (h === '患者姓名') ci.patient = i;
  if (h.includes('合同创建') && h.includes('日期')) ci.od = i;
  if (h.includes('实际回输') && h.includes('结束时间')) ci.re = i;
});

var records = [];
var unmatchedWarnings = [];

for (var i = 2; i < bsRows.length; i++) {
  var row = bsRows[i]; if (!row) continue;
  var od = excelToDate(row[ci.od]), re = parseDt(row[ci.re]);
  var code1 = String(row[ci.org] || '').trim();
  var code2 = String(row[5] || '').trim();
  var m = null;

  if (code1 && masterMap[code1]) {
    m = masterMap[code1];
  } else if (code2 && masterMap[code2]) {
    m = masterMap[code2];
  }

  var hosp, prov;
  if (m) {
    hosp = m.name; prov = m.prov;
  } else {
    var fallbackHosp = String(row[ci.orgName] || row[6] || '').trim();
    if (code1 || code2) {
      unmatchedWarnings.push({
        row: i, code1: code1, code2: code2,
        rawHosp: fallbackHosp || row[6] || ''
      });
    }
    hosp = fallbackHosp; prov = '';
  }

  if (!hosp) continue;
  var patient = maskName(row[ci.patient]);
  var isSG = (prov === '新加坡' || prov.includes('新加坡'));
  records.push({ hosp, prov: isSG ? '' : prov, od, re, noMap: isSG, patient: patient });
}

if (unmatchedWarnings.length > 0) {
  console.error('\n========================================');
  console.error('⚠️  警告：发现 ' + unmatchedWarnings.length + ' 条记录无法匹配主数据！');
  var seen = new Set();
  unmatchedWarnings.forEach(function(w) {
    var key = w.code1 || w.code2;
    if (!seen.has(key)) {
      seen.add(key);
      console.error('  · 编码: ' + key + '  原始医院名: ' + (w.rawHosp || '未知'));
    }
  });
  console.error('请将以上编码添加至 masterdata.xlsx 后重新运行！');
  console.error('========================================\n');
  process.exit(1);
}

function inRange(d, f, t) { return d && d >= f && d <= t; }

var DP = toLocal(new Date());
var Y = DP.slice(0, 4);
var dpM = DP.slice(5, 7);
var dpM0 = parseInt(dpM);

var ytdO = records.filter(r => inRange(r.od, Y + '-01-01', DP)).length;
var ytdR = records.filter(r => inRange(r.re, Y + '-01-01', DP)).length;
var mtdO = records.filter(r => inRange(r.od, Y + '-' + dpM + '-01', DP)).length;
var mtdR = records.filter(r => inRange(r.re, Y + '-' + dpM + '-01', DP)).length;

var monO = [0,0,0,0,0,0,0,0,0,0,0,0], monR = [0,0,0,0,0,0,0,0,0,0,0,0];
records.forEach(r => {
  if (r.od && r.od >= Y + '-01-01') { var mi = parseInt(r.od.slice(5,7)) - 1; if (mi < 12) monO[mi]++; }
  if (r.re && r.re >= Y + '-01-01') { var mi = parseInt(r.re.slice(5,7)) - 1; if (mi < 12) monR[mi]++; }
});

var dpDate = new Date(DP.replace(/-/g, '/'));
var p7 = [];
for (var d = 6; d >= 0; d--) {
  var dt = new Date(dpDate); dt.setDate(dt.getDate() - d);
  var ds = toLocal(dt);
  var oMap = {}, rMap = {};
  records.forEach(r => {
    if (r.od === ds) {
      var key = r.hosp + '|||' + r.patient;
      oMap[key] = (oMap[key]||0)+1;
    }
    if (r.re === ds) {
      var key = r.hosp + '|||' + r.patient;
      rMap[key] = (rMap[key]||0)+1;
    }
  });
  p7.push({ date: ds, orders: oMap, reinfusion: rMap });
}

var provGeoMap = { '香港':'香港特别行政区','澳门':'澳门特别行政区','台湾':'台湾省','新加坡':'' };
var provData = {};
records.forEach(r => {
  if (r.noMap) return;
  if (r.od && r.od >= Y+'-01-01' && r.od <= DP) {
    var p = provGeoMap[r.prov] !== undefined ? provGeoMap[r.prov] : r.prov;
    if (p) provData[p] = (provData[p]||0)+1;
  }
});

var ALL_P = ['北京市','天津市','上海市','重庆市','河北省','山西省','辽宁省','吉林省','黑龙江省','江苏省','浙江省','安徽省','福建省','江西省','山东省','河南省','湖北省','湖南省','广东省','广西壮族自治区','海南省','四川省','贵州省','云南省','西藏自治区','陕西省','甘肃省','青海省','宁夏回族自治区','新疆维吾尔自治区','内蒙古自治区','台湾省','香港特别行政区','澳门特别行政区'];
var mapJson = ALL_P.map(p => ({ name: p, value: provData[p]||0 }));
var provRank5 = Object.entries(provData).sort((a,b) => b[1]-a[1]).slice(0, 5);

p7.reverse();

var p7totalO = 0, p7totalR = 0;
p7.forEach(d => {
  p7totalO += Object.values(d.orders).reduce((a,b)=>a+b,0);
  p7totalR += Object.values(d.reinfusion).reduce((a,b)=>a+b,0);
});

var topProvNames = provRank5.slice(0,3).map(e =>
  e[0].replace('北京市','北京').replace('天津市','天津').replace('上海市','上海')
     .replace('特别行政区','').replace('壮族自治区','').replace('省','').replace('市','')
);

var todayO = records.filter(r => r.od === DP).length;
var todayR = records.filter(r => r.re === DP).length;

var summary = {
  DP, Y, dpM, dpM0,
  ytdO, ytdR, mtdO, mtdR,
  todayO, todayR,
  p7totalO, p7totalR,
  topProvNames,
  monO, monR,
  mapJson, provRank5,
  p7
};

// ============================================================
// 输出 data-inline.js（供网页内置引用）
// ============================================================
var outPath = path.join(__dirname, 'data-inline.js');
var outContent = '/** 内置数据 — 由 build-data.js 自动生成 */\nwindow.BUILTIN_DATA = ' + JSON.stringify(summary, null, 2) + ';\n';
fs.writeFileSync(outPath, outContent, 'utf-8');

console.log('\n✅ 内置数据已生成: ' + outPath);
console.log('YTD下单:', ytdO, ' YTD回输:', ytdR);
console.log('MTD下单:', mtdO, ' MTD回输:', mtdR);
console.log('当日('+DP+'): 下单'+todayO+' 回输'+todayR);
console.log('过去7天: 下单'+p7totalO+' 回输'+p7totalR);
console.log('Top5省份:', provRank5.map(e=>e[0]+'('+e[1]+')').join(', '));
console.log('\n下一步：将 data-inline.js 提交到 GitHub，网站会自动展示最新数据。');
