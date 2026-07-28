/**
 * 福可苏 FUCASO 销量日报 · 网页版 Dashboard
 * 浏览器端数据处理 + 可视化渲染
 */

// ============================================================
// 状态管理
// ============================================================
const state = {
  files: { bsOrder: null, masterdata: null },
  records: [],
  summary: {},
  charts: {}
};

// ============================================================
// 登录验证配置
// ============================================================
const ACCESS_PASSWORD = 'fucaso2026';  // ⚠️ 修改此处为你要设置的密码
const LOGIN_KEY = 'fucaso_login_token';

// ============================================================
// UI 工具函数
// ============================================================
function showToast(message, type = 'info') {
  const toast = document.getElementById('toast');
  toast.textContent = message;
  toast.className = 'toast ' + type;
  requestAnimationFrame(() => toast.classList.add('show'));
  setTimeout(() => toast.classList.remove('show'), 4000);
}

function setLoading(text, active = true) {
  const overlay = document.getElementById('loadingOverlay');
  document.getElementById('loadingText').textContent = text;
  overlay.classList.toggle('active', active);
}

function switchPage(page) {
  document.getElementById('uploadPage').style.display = page === 'upload' ? 'flex' : 'none';
  document.getElementById('dashboardWrap').classList.toggle('active', page === 'dashboard');
}

function updateFileList() {
  const list = document.getElementById('fileList');
  const btn = document.getElementById('btnGenerate');
  const parts = [];

  if (state.files.bsOrder) {
    parts.push(`<div class="file-tag"><span>📄 ${escapeHtml(state.files.bsOrder.name)}</span><span class="del" data-type="bsOrder">✕</span></div>`);
  }
  if (state.files.masterdata) {
    parts.push(`<div class="file-tag"><span>📄 ${escapeHtml(state.files.masterdata.name)}</span><span class="del" data-type="masterdata">✕</span></div>`);
  }

  list.innerHTML = parts.join('');

  list.querySelectorAll('.del').forEach(el => {
    el.addEventListener('click', e => {
      e.stopPropagation();
      const type = el.dataset.type;
      state.files[type] = null;
      updateFileList();
    });
  });

  const ready = state.files.bsOrder && state.files.masterdata;
  btn.disabled = !ready;
  btn.textContent = ready ? '🚀 生成销量日报' : '⏳ 等待上传文件…';
}

function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

// ============================================================
// 登录验证
// ============================================================
function initLogin() {
  const overlay = document.getElementById('loginOverlay');
  const input = document.getElementById('loginInput');
  const btn = document.getElementById('btnLogin');
  const error = document.getElementById('loginError');

  if (sessionStorage.getItem(LOGIN_KEY) === '1') {
    overlay.classList.add('hidden');
    return true;
  }

  setTimeout(() => input.focus(), 300);

  function doLogin() {
    const val = input.value.trim();
    if (val === ACCESS_PASSWORD) {
      sessionStorage.setItem(LOGIN_KEY, '1');
      overlay.classList.add('hidden');
      error.textContent = '';
      initDashboard();
    } else {
      error.textContent = '❌ 密码错误，请重试';
      input.value = '';
      input.focus();
    }
  }

  btn.addEventListener('click', doLogin);
  input.addEventListener('keydown', e => {
    if (e.key === 'Enter') doLogin();
  });

  return false;
}

// ============================================================
// 文件上传处理
// ============================================================
function initUpload() {
  const zone = document.getElementById('uploadZone');
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = '.xlsx,.xls';
  input.multiple = true;
  input.style.display = 'none';
  document.body.appendChild(input);

  zone.addEventListener('click', () => input.click());

  zone.addEventListener('dragover', e => {
    e.preventDefault();
    zone.classList.add('dragover');
  });
  zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.classList.remove('dragover');
    handleFiles(e.dataTransfer.files);
  });

  input.addEventListener('change', () => handleFiles(input.files));
}

function handleFiles(fileList) {
  for (const file of fileList) {
    const name = file.name.toLowerCase();
    if (name.startsWith('bs_order') && !name.startsWith('~$')) {
      state.files.bsOrder = file;
    } else if (name === 'masterdata.xlsx' || name === 'masterdata.xls') {
      state.files.masterdata = file;
    } else if (name.includes('masterdata')) {
      state.files.masterdata = file;
    }
  }
  updateFileList();
}

// ============================================================
// 核心数据处理
// ============================================================

function toLocal(d) {
  return d.getFullYear() + '-' +
    String(d.getMonth() + 1).padStart(2, '0') + '-' +
    String(d.getDate()).padStart(2, '0');
}

function excelToDate(v) {
  if (typeof v === 'number') return toLocal(new Date((v - 25569) * 86400 * 1000));
  if (typeof v === 'string') {
    var m = v.match(/(\d{4}-\d{2}-\d{2})/);
    return m ? m[1] : null;
  }
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

function inRange(d, f, t) { return d && d >= f && d <= t; }
function topN(obj, n) { return Object.entries(obj).sort((a, b) => b[1] - a[1]).slice(0, n); }
function fmtNum(n) { return n.toLocaleString(); }

// ============================================================
// Excel 读取
// ============================================================
function readExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        resolve(workbook);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// ============================================================
// 主处理流程
// ============================================================
async function processData() {
  setLoading('正在读取 Excel 文件…');

  try {
    const [bsWb, mdWb] = await Promise.all([
      readExcelFile(state.files.bsOrder),
      readExcelFile(state.files.masterdata)
    ]);

    setLoading('正在解析数据…');

    const bsSheet = bsWb.Sheets[bsWb.SheetNames[0]];
    const bsRows = XLSX.utils.sheet_to_json(bsSheet, { header: 1 });

    const mdSheet = mdWb.Sheets[mdWb.SheetNames[0]];
    const mdRows = XLSX.utils.sheet_to_json(mdSheet);

    setLoading('正在关联主数据…');

    var masterMap = {};
    mdRows.forEach(r => {
      var c = String(r['细胞追溯系统代码'] || '').trim();
      if (c) masterMap[c] = {
        name: String(r['标准医院名称'] || '').trim(),
        prov: String(r['省份'] || '').trim()
      };
    });

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
      records.push({ hosp, prov: isSG ? '' : prov, od, re, ap, qa, noMap: isSG, patient: patient });
    }

    if (unmatchedWarnings.length > 0) {
      var seen = new Set();
      var msg = '⚠️ 发现 ' + unmatchedWarnings.length + ' 条记录无法匹配主数据：\n';
      unmatchedWarnings.forEach(function(w) {
        var key = w.code1 || w.code2;
        if (!seen.has(key)) {
          seen.add(key);
          msg += '  · 编码: ' + key + '  原始医院名: ' + (w.rawHosp || '未知') + '\n';
        }
      });
      msg += '\n请将以上编码添加至 masterdata.xlsx 后重新上传！';
      throw new Error(msg);
    }

    setLoading('正在计算指标…');

    var DP = toLocal(new Date());
    var Y = DP.slice(0, 4);
    console.log('数据截止日:', DP);

    var ytdO = records.filter(r => inRange(r.od, Y + '-01-01', DP)).length;
    var ytdR = records.filter(r => inRange(r.re, Y + '-01-01', DP)).length;

    var dpM = DP.slice(5, 7);
    var dpM0 = parseInt(dpM);
    var mtdO = records.filter(r => inRange(r.od, Y + '-' + dpM + '-01', DP)).length;
    var mtdR = records.filter(r => inRange(r.re, Y + '-' + dpM + '-01', DP)).length;

    var monO = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
    var monR = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
    records.forEach(r => {
      if (r.od && r.od >= Y + '-01-01') { var mi = parseInt(r.od.slice(5, 7)) - 1; if (mi < 12) monO[mi]++; }
      if (r.re && r.re >= Y + '-01-01') { var mi = parseInt(r.re.slice(5, 7)) - 1; if (mi < 12) monR[mi]++; }
    });

    var WDAYS = ['周日', '周一', '周二', '周三', '周四', '周五', '周六'];
    var dpDate = new Date(DP.replace(/-/g, '/'));
    var p7 = [];
    for (var d = 6; d >= 0; d--) {
      var dt = new Date(dpDate); dt.setDate(dt.getDate() - d);
      var ds = toLocal(dt);
      var oMap = {}, rMap = {}, aMap = {}, qMap = {};
      records.forEach(r => {
        if (r.od === ds) {
          var key = r.hosp + '|||' + r.patient;
          oMap[key] = (oMap[key] || 0) + 1;
        }
        if (r.re === ds) {
          var key = r.hosp + '|||' + r.patient;
          rMap[key] = (rMap[key] || 0) + 1;
        }
        if (r.ap === ds) {
          var key = r.hosp + '|||' + r.patient;
          aMap[key] = (aMap[key] || 0) + 1;
        }
        if (r.qa === ds) {
          var key = r.hosp + '|||' + r.patient;
          qMap[key] = (qMap[key] || 0) + 1;
        }
      });
      p7.push({ date: ds, orders: oMap, reinfusion: rMap, apheresis: aMap, quality: qMap });
    }

    var provGeoMap = { '香港': '香港特别行政区', '澳门': '澳门特别行政区', '台湾': '台湾省', '新加坡': '' };
    var provData = {};
    records.forEach(r => {
      if (r.noMap) return;
      if (r.od && r.od >= Y + '-01-01' && r.od <= DP) {
        var p = provGeoMap[r.prov] !== undefined ? provGeoMap[r.prov] : r.prov;
        if (p) provData[p] = (provData[p] || 0) + 1;
      }
    });

    var ALL_P = ['北京市', '天津市', '上海市', '重庆市', '河北省', '山西省', '辽宁省', '吉林省', '黑龙江省', '江苏省', '浙江省', '安徽省', '福建省', '江西省', '山东省', '河南省', '湖北省', '湖南省', '广东省', '广西壮族自治区', '海南省', '四川省', '贵州省', '云南省', '西藏自治区', '陕西省', '甘肃省', '青海省', '宁夏回族自治区', '新疆维吾尔自治区', '内蒙古自治区', '台湾省', '香港特别行政区', '澳门特别行政区'];
    var mapJson = ALL_P.map(p => ({ name: p, value: provData[p] || 0 }));
    var provRank5 = Object.entries(provData).sort((a, b) => b[1] - a[1]).slice(0, 5);

    p7.reverse();

    var p7totalO = 0, p7totalR = 0, p7totalA = 0, p7totalQ = 0;
    p7.forEach(d => {
      p7totalO += Object.values(d.orders).reduce((a, b) => a + b, 0);
      p7totalR += Object.values(d.reinfusion).reduce((a, b) => a + b, 0);
      p7totalA += Object.values(d.apheresis).reduce((a, b) => a + b, 0);
      p7totalQ += Object.values(d.quality).reduce((a, b) => a + b, 0);
    });

    var topProvNames = provRank5.slice(0, 3).map(e =>
      e[0].replace('北京市', '北京').replace('天津市', '天津').replace('上海市', '上海')
        .replace('特别行政区', '').replace('壮族自治区', '').replace('省', '').replace('市', '')
    );

    var todayO = records.filter(r => r.od === DP).length;
    var todayR = records.filter(r => r.re === DP).length;
    var todayA = records.filter(r => r.ap === DP).length;
    var todayQ = records.filter(r => r.qa === DP).length;

    state.records = records;
    state.summary = {
      DP, Y, dpM, dpM0,
      ytdO, ytdR, mtdO, mtdR,
      todayO, todayR, todayA, todayQ,
      p7totalO, p7totalR, p7totalA, p7totalQ,
      topProvNames,
      monO, monR,
      mapJson, provRank5,
      p7
    };

    setLoading('正在渲染看板…');
    renderDashboard();

    setLoading('', false);
    switchPage('dashboard');
    showToast('✅ 日报生成成功！', 'success');

  } catch (err) {
    setLoading('', false);
    console.error(err);
    showToast(err.message || '处理失败，请检查文件格式', 'error');
  }
}

// ============================================================
// Dashboard 渲染
// ============================================================
function renderDashboard() {
  const s = state.summary;
  const container = document.getElementById('posterContent');

  var tableRows = '';
  s.p7.forEach(day => {
    var dObj = new Date(day.date.replace(/-/g, '/'));
    var mmdd = day.date.slice(5);
    var wd = ['周日', '周一', '周二', '周三', '周四', '周五', '周六'][dObj.getDay()];
    var oTot = Object.values(day.orders).reduce((a, b) => a + b, 0);
    var rTot = Object.values(day.reinfusion).reduce((a, b) => a + b, 0);
    var aTot = Object.values(day.apheresis).reduce((a, b) => a + b, 0);
    var qTot = Object.values(day.quality).reduce((a, b) => a + b, 0);

    var oLi = topN(day.orders, 12).map(e => {
      var parts = e[0].split('|||');
      return '<li><span class="hn">' + parts[0] + '</span><span class="hp">' + parts[1] + '</span><span class="hc o">' + e[1] + '单</span></li>';
    }).join('');
    var rLi = topN(day.reinfusion, 12).map(e => {
      var parts = e[0].split('|||');
      return '<li><span class="hn">' + parts[0] + '</span><span class="hp">' + parts[1] + '</span><span class="hc r">' + e[1] + '单</span></li>';
    }).join('');
    var aLi = topN(day.apheresis, 12).map(e => {
      var parts = e[0].split('|||');
      return '<li><span class="hn">' + parts[0] + '</span><span class="hp">' + parts[1] + '</span><span class="hc a">' + e[1] + '单</span></li>';
    }).join('');
    var qLi = topN(day.quality, 12).map(e => {
      var parts = e[0].split('|||');
      return '<li><span class="hn">' + parts[0] + '</span><span class="hp">' + parts[1] + '</span><span class="hc q">' + e[1] + '单</span></li>';
    }).join('');
    if (!oLi) oLi = '<li><span class="dim">暂无数据</span></li>';
    if (!rLi) rLi = '<li><span class="dim">暂无数据</span></li>';
    if (!aLi) aLi = '<li><span class="dim">暂无数据</span></li>';
    if (!qLi) qLi = '<li><span class="dim">暂无数据</span></li>';

    var dayTotal = '<span class="d-o">下单' + oTot + '</span><span class="d-r">回输' + rTot + '</span><span class="d-a">单采' + aTot + '</span><span class="d-q">放行' + qTot + '</span>';
    tableRows += '<tr><td class="cell-date"><div class="date-tl">' + mmdd + '</div><div class="date-wd">' + wd + '</div><div class="date-sum">' + dayTotal + '</div></td><td><ul class="hlist">' + oLi + '</ul></td><td><ul class="hlist">' + rLi + '</ul></td><td><ul class="hlist">' + aLi + '</ul></td><td><ul class="hlist">' + qLi + '</ul></td></tr>\n';
  });

  var rankHtml = '';
  s.provRank5.forEach(function(e, i) {
    var rankColors = ['gold', 'silver', 'bronze', 'normal', 'normal'];
    var shortName = e[0].replace('北京市', '北京').replace('天津市', '天津').replace('上海市', '上海').replace('重庆市', '重庆')
      .replace('特别行政区', '').replace('壮族自治区', '广西').replace('维吾尔自治区', '新疆').replace('回族自治区', '宁夏').replace('自治区', '').replace('省', '');
    var maxV = s.provRank5[0][1];
    var barW = Math.round(e[1] / maxV * 80);
    rankHtml += '<div class="rank-item"><div class="r-idx ' + rankColors[i] + '">' + (i + 1) + '</div><div class="r-name">' + shortName + '</div><div class="r-bar" style="width:' + barW + 'px;background:var(--blue);"></div><div class="r-val">' + e[1] + '</div></div>';
  });

  container.innerHTML = `
    <div class="header">
      <div class="h-left">
        <div class="logo">
          <span style="font-size:14px;font-weight:700;color:#0d1b2e;letter-spacing:2px;">IASO 驯鹿生物</span>
        </div>
        <div class="h-title">
          <div class="cn">福可苏 FUCASO · 销量日报</div>
          <div class="en">FUCASO DAILY SALES REPORT</div>
        </div>
      </div>
      <div class="h-right">
        <div class="badge">${s.Y}年日报</div>
        <span class="ts">数据截止 ${s.DP}</span>
        <div class="h-summary">
          <b>YTD下单</b> ${fmtNum(s.ytdO)} 单，<b>回输</b> ${fmtNum(s.ytdR)} 单；
          <b>${s.dpM0}月MTD</b> 下单 ${fmtNum(s.mtdO)} 单，回输 ${fmtNum(s.mtdR)} 单
        </div>
      </div>
    </div>

    <div class="kpi-row">
      <div class="kpi-card oy">
        <div class="section-label">📋 全年累计</div>
        <div class="kh"><span class="kl">${s.Y} YTD · 下单销量</span></div>
        <div class="kv">${fmtNum(s.ytdO)}</div>
        <div class="ks">合同创建日期统计</div>
      </div>
      <div class="kpi-card ry">
        <div class="section-label">💉 全年累计</div>
        <div class="kh"><span class="kl">${s.Y} YTD · 回输销量</span></div>
        <div class="kv">${fmtNum(s.ytdR)}</div>
        <div class="ks">实际回输结束时间统计</div>
      </div>
      <div class="kpi-card om">
        <div class="section-label">当月进度</div>
        <div class="kh"><span class="kl">${s.dpM0}月 MTD · 下单</span></div>
        <div class="kv">${fmtNum(s.mtdO)}</div>
        <div class="ks">MTD</div>
      </div>
      <div class="kpi-card rm">
        <div class="section-label">当月进度</div>
        <div class="kh"><span class="kl">${s.dpM0}月 MTD · 回输</span></div>
        <div class="kv">${fmtNum(s.mtdR)}</div>
        <div class="ks">MTD</div>
      </div>
    </div>

    <div class="insight-bar">
      <div class="dot"></div>
      <b>本周观察：</b>过去7天累计下单 <b>${s.p7totalO}</b> 单，回输 <b>${s.p7totalR}</b> 单，单采 <b>${s.p7totalA}</b> 单，放行 <b>${s.p7totalQ}</b> 单；${s.topProvNames.join('、')}为主要贡献区域。
    </div>

    <div class="section-block">
      <div class="sec-head">
        <div class="line"></div>
        <span class="ttl">过去七天 · 下单 & 回输 & 单采 & 放行明细</span>
        <span class="sub">${s.p7[6].date} → ${s.p7[0].date}</span>
      </div>
      <table class="timeline-table">
        <thead>
          <tr><th style="width:10%">日期</th><th style="width:22.5%">下单明细</th><th style="width:22.5%">回输明细</th><th style="width:22.5%">单采明细</th><th style="width:22.5%">放行明细</th></tr>
        </thead>
        <tbody>
          ${tableRows}
        </tbody>
      </table>
    </div>

    <div class="section-block">
      <div class="sec-head">
        <div class="line"></div>
        <span class="ttl">${s.Y}年月度销量趋势｜下单 vs 回输</span>
        <span class="sub">${s.dpM0}月为MTD数据</span>
      </div>
      <div class="legend-row">
        <span><span class="dot" style="background:var(--blue)"></span>下单销量</span>
        <span><span class="dot" style="background:var(--green)"></span>回输销量</span>
        <span style="margin-left:16px;font-size:11px;color:var(--t3)">虚线=MTD</span>
      </div>
      <div class="chart-wrap" id="monthlyBar"></div>
    </div>

    <div class="section-block">
      <div class="sec-head">
        <div class="line"></div>
        <span class="ttl">${s.Y}年 各省份下单数量热力分布</span>
        <span class="sub">单位：单</span>
      </div>
      <div class="map-row">
        <div class="map-wrap" id="chinaMap"></div>
        <div class="rank-panel">
          <div class="r-title">📋 下单销量 TOP 5</div>
          ${rankHtml}
        </div>
      </div>
    </div>

    <div class="footer">
      <span>福可苏 FUCASO · 销量数据报告</span>
      <span>Generated by Data Command Center</span>
    </div>
  `;

  setTimeout(() => {
    initMonthlyBarChart();
    initChinaMapChart();
  }, 100);
}

// ============================================================
// ECharts 图表初始化
// ============================================================
function initMonthlyBarChart() {
  const s = state.summary;
  const el = document.getElementById('monthlyBar');
  if (!el) return;

  if (state.charts.bar) {
    state.charts.bar.dispose();
  }

  var c = echarts.init(el, null, { devicePixelRatio: 2 });
  var currM = s.dpM0;
  var oData = [], rData = [];

  for (var i = 0; i < 12; i++) {
    var isMTD = i === currM - 1;
    var isFuture = i >= currM;
    oData.push({
      value: isFuture ? null : s.monO[i],
      itemStyle: isFuture ? { color: 'transparent' } : isMTD ?
        { color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [{ offset: 0, color: '#6db3ff' }, { offset: 1, color: '#2563eb' }]), opacity: 0.7 } :
        { color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [{ offset: 0, color: '#6db3ff' }, { offset: 1, color: '#2563eb' }]) }
    });
    rData.push({
      value: isFuture ? null : s.monR[i],
      itemStyle: isFuture ? { color: 'transparent' } : isMTD ?
        { color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [{ offset: 0, color: '#4ddfc0' }, { offset: 1, color: '#059669' }]), opacity: 0.7 } :
        { color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [{ offset: 0, color: '#4ddfc0' }, { offset: 1, color: '#059669' }]) }
    });
  }

  var MONTHS = ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月'];

  c.setOption({
    grid: { left: 14, right: 24, top: 36, bottom: 28 },
    xAxis: {
      type: 'category',
      data: MONTHS,
      axisLine: { lineStyle: { color: 'rgba(255,255,255,.06)' } },
      axisTick: { show: false },
      axisLabel: { color: '#5a6688', fontSize: 12, margin: 8 }
    },
    yAxis: {
      type: 'value',
      min: 0,
      showMinLabel: false,
      splitLine: { lineStyle: { color: 'rgba(255,255,255,.04)', type: 'dashed' } },
      axisLabel: { color: '#5a6688', fontSize: 12 }
    },
    series: [
      {
        name: '下单销量',
        type: 'bar',
        data: oData,
        barWidth: 22,
        barGap: '35%',
        itemStyle: { borderRadius: [4, 4, 0, 0] },
        label: {
          show: true,
          position: 'top',
          color: '#93c5fd',
          fontSize: 13,
          fontWeight: 600,
          formatter: function(p) {
            var v = p.data ? p.data.value : 0;
            return v ? String(v) : '';
          }
        },
        markLine: {
          silent: true,
          symbol: 'none',
          label: { show: false },
          lineStyle: { color: '#f59e0b', type: 'dashed', width: 1 },
          data: [{
            xAxis: currM + '月',
            label: { show: true, formatter: 'MTD', color: '#f59e0b', fontSize: 12, fontWeight: 600, position: 'end' }
          }]
        }
      },
      {
        name: '回输销量',
        type: 'bar',
        data: rData,
        barWidth: 22,
        itemStyle: { borderRadius: [4, 4, 0, 0] },
        label: {
          show: true,
          position: 'top',
          color: '#6ee7b7',
          fontSize: 13,
          fontWeight: 600,
          formatter: function(p) {
            var v = p.data ? p.data.value : 0;
            return v ? String(v) : '';
          }
        }
      }
    ]
  });

  state.charts.bar = c;
}

function initChinaMapChart() {
  const s = state.summary;
  const el = document.getElementById('chinaMap');
  if (!el) return;

  if (state.charts.map) {
    state.charts.map.dispose();
  }

  var c = echarts.init(el, null, { devicePixelRatio: 2 });

  fetch('https://geo.datav.aliyun.com/areas_v3/bound/100000_full.json')
    .then(r => r.json())
    .then(g => {
      echarts.registerMap('china', g);
      var pal = ['#0d2845', '#144d6e', '#1a7890', '#22a0a8', '#38c9b0', '#65e8c8', '#a0ffe0'];
      var vals = s.mapJson.map(x => x.value);
      var maxV = Math.max.apply(null, vals) || 1;
      var sd = s.mapJson.map(x => {
        var t = x.value / maxV;
        var idx = Math.min(Math.floor(t * (pal.length - 1)), pal.length - 1);
        return { name: x.name, value: x.value, itemStyle: { areaColor: x.value === 0 ? '#0a1226' : pal[idx] } };
      });

      c.setOption({
        series: [{
          type: 'map',
          map: 'china',
          roam: false,
          zoom: 1.3,
          center: [105, 37],
          label: {
            show: true,
            color: '#b0c8e0',
            fontSize: 9,
            formatter: function(p) {
              var raw = p.name;
              var name = raw;
              if (raw === '内蒙古自治区') name = '内蒙古';
              if (raw === '广西壮族自治区') name = '广西';
              if (raw === '西藏自治区') name = '西藏';
              if (raw === '宁夏回族自治区') name = '宁夏';
              if (raw === '新疆维吾尔自治区') name = '新疆';
              if (raw === '香港特别行政区') name = '香港';
              if (raw === '澳门特别行政区') name = '澳门';
              if (raw === '澳门特别行政区') return '';
              if (!p.data || !p.data.value) return name;
              var val = p.data.value;
              if (val === 0) return name;
              if (raw === '北京市') return '北京  ' + val + '\n\n';
              if (raw === '天津市') return ' \n'.repeat(3) + '天津  ' + val;
              if (raw === '上海市') return '上海  ' + val;
              if (raw === '广东省') return '广东  ' + val;
              if (raw === '香港特别行政区') return '香港  ' + val;
              return name + '\n' + val;
            }
          },
          emphasis: {
            label: { show: true, color: '#fff', fontSize: 14, fontWeight: 'bold' },
            itemStyle: { areaColor: '#f59e0b', borderColor: '#fff', borderWidth: 1.5 }
          },
          itemStyle: {
            borderColor: 'rgba(255,255,255,.25)',
            borderWidth: 1,
            areaColor: '#0a1226'
          },
          data: sd
        }]
      });
    })
    .catch(e => {
      el.innerHTML = '<div style="display:flex;align-items:center;justify-content:center;height:100%;color:#5a6688;">地图加载失败，请检查网络连接</div>';
      console.error('地图加载失败:', e);
    });

  state.charts.map = c;
}

// ============================================================
// PNG 下载
// ============================================================
async function downloadPNG() {
  const btn = document.getElementById('btnDownload');
  const originalText = btn.textContent;
  btn.disabled = true;
  btn.textContent = '⏳ 生成中…';

  try {
    const canvas = await html2canvas(document.getElementById('poster'), {
      scale: 2,
      useCORS: true,
      backgroundColor: '#06111f',
      allowTaint: false,
      logging: false
    });

    var d = new Date();
    var ds = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
    var lnk = document.createElement('a');
    lnk.download = '福可苏_销量日报_' + ds + '.png';
    lnk.href = canvas.toDataURL('image/png');
    lnk.click();

    btn.textContent = '✅ 下载完成';
    setTimeout(() => { btn.textContent = originalText; btn.disabled = false; }, 2000);
  } catch (e) {
    console.error(e);
    btn.textContent = '❌ 截图失败';
    setTimeout(() => { btn.textContent = originalText; btn.disabled = false; }, 2000);
  }
}

// ============================================================
// 窗口 resize
// ============================================================
function handleResize() {
  if (state.charts.bar) state.charts.bar.resize();
  if (state.charts.map) state.charts.map.resize();
}

// ============================================================
// Dashboard 初始化（双模式：内置数据 或 上传模式）
// ============================================================
function initDashboard() {
  // 检查是否有内置数据（由 build-data.js 生成）
  if (window.BUILTIN_DATA) {
    console.log('✅ 检测到内置数据，直接渲染 Dashboard');
    state.summary = window.BUILTIN_DATA;
    switchPage('dashboard');
    renderDashboard();
    document.getElementById('btnDownload').addEventListener('click', downloadPNG);
    document.getElementById('btnReupload').addEventListener('click', () => {
      state.files = { bsOrder: null, masterdata: null };
      state.records = [];
      state.summary = {};
      if (state.charts.bar) { state.charts.bar.dispose(); state.charts.bar = null; }
      if (state.charts.map) { state.charts.map.dispose(); state.charts.map = null; }
      document.getElementById('posterContent').innerHTML = '';
      initUploadMode();
      switchPage('upload');
    });
    initAdminToggle();
    window.addEventListener('resize', handleResize);
    return;
  }

  // 无内置数据，显示上传页面
  console.log('ℹ️ 无内置数据，显示上传页面');
  initUploadMode();
}

function initAdminToggle() {
  var actionsBar = document.querySelector('.actions-bar');
  if (actionsBar) {
    var adminBtn = document.createElement('button');
    adminBtn.className = 'btn btn-secondary';
    adminBtn.textContent = '🔧 管理员更新数据';
    adminBtn.onclick = function() {
      if (confirm('切换到管理员上传模式？当前 Dashboard 数据将被替换。')) {
        state.summary = {};
        if (state.charts.bar) { state.charts.bar.dispose(); state.charts.bar = null; }
        if (state.charts.map) { state.charts.map.dispose(); state.charts.map = null; }
        document.getElementById('posterContent').innerHTML = '';
        initUploadMode();
        switchPage('upload');
      }
    };
    actionsBar.appendChild(adminBtn);
  }
}

function initUploadMode() {
  initUpload();
  document.getElementById('btnGenerate').addEventListener('click', processData);
  document.getElementById('btnDownload').addEventListener('click', downloadPNG);
  document.getElementById('btnReupload').addEventListener('click', () => {
    state.files = { bsOrder: null, masterdata: null };
    state.records = [];
    state.summary = {};
    if (state.charts.bar) { state.charts.bar.dispose(); state.charts.bar = null; }
    if (state.charts.map) { state.charts.map.dispose(); state.charts.map = null; }
    updateFileList();
    switchPage('upload');
  });
  window.addEventListener('resize', handleResize);
}

// ============================================================
// 入口
// ============================================================
document.addEventListener('DOMContentLoaded', () => {
  const loggedIn = initLogin();
  if (loggedIn) {
    initDashboard();
  }
});
