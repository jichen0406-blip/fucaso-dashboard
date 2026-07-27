#!/usr/bin/env node
/**
 * deploy.js — 一键更新数据并部署到 GitHub
 *
 * 用法：node deploy.js
 *
 * 功能：
 * 1. 检查 rawdata 中的 Excel 文件
 * 2. 运行 build-data.js 生成 data-inline.js
 * 3. 检查数据是否有变化
 * 4. 自动 git add → commit → push
 * 5. 等待 GitHub Pages 部署完成
 */

const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const ROOT = __dirname;
// ⚠️ 修改此处为你的 rawdata 实际路径
const RAWDATA_DIR = 'D:\\VS Code\\fucaso-dashboard\\rawdata';

// ─── 彩色输出工具 ───
const C = {
  reset: '\x1b[0m',
  bold: '\x1b[1m',
  green: '\x1b[32m',
  red: '\x1b[31m',
  yellow: '\x1b[33m',
  blue: '\x1b[34m',
  cyan: '\x1b[36m',
  gray: '\x1b[90m'
};
function log(msg, color) {
  console.log((color || '') + msg + C.reset);
}
function info(msg) { log('ℹ️  ' + msg, C.blue); }
function ok(msg) { log('✅ ' + msg, C.green); }
function warn(msg) { log('⚠️  ' + msg, C.yellow); }
function err(msg) { log('❌ ' + msg, C.red); }
function step(n, msg) {
  log('\n' + C.bold + C.cyan + '═══ 步骤 ' + n + '：' + msg + ' ═══' + C.reset);
}

// ─── 执行命令（带错误处理）───
function run(cmd, opts) {
  opts = opts || {};
  try {
    const out = execSync(cmd, {
      cwd: ROOT,
      encoding: 'utf-8',
      stdio: opts.silent ? 'pipe' : 'inherit',
      timeout: opts.timeout || 60000
    });
    return out ? out.trim() : '';
  } catch (e) {
    if (opts.ignoreError) return '';
    throw e;
  }
}

// ─── 主流程 ───
function main() {
  log('\n' + C.bold + '╔══════════════════════════════════════════╗' + C.reset);
  log(C.bold + '║  福可苏 FUCASO · 一键数据更新 & 自动部署  ║' + C.reset);
  log(C.bold + '╚══════════════════════════════════════════╝' + C.reset);

  // ─── 步骤 1：检查环境 ───
  step(1, '检查环境');

  try {
    info('Node.js 版本: ' + run('node --version', { silent: true }));
  } catch (e) {
    err('未找到 Node.js，请先安装');
    process.exit(1);
  }

  try {
    info('Git 版本: ' + run('git --version', { silent: true }));
  } catch (e) {
    err('未找到 Git，请先安装');
    process.exit(1);
  }

  if (!fs.existsSync(RAWDATA_DIR)) {
    err('rawdata 目录不存在: ' + RAWDATA_DIR);
    info('请将 bs_order 和 masterdata Excel 文件放入上述目录');
    process.exit(1);
  }

  const rawFiles = fs.readdirSync(RAWDATA_DIR);
  const bsFiles = rawFiles.filter(function(f) {
    return f.startsWith('bs_order') && !f.startsWith('~$');
  });
  const mdFiles = rawFiles.filter(function(f) {
    return f.toLowerCase().includes('masterdata') && !f.startsWith('~$');
  });

  if (bsFiles.length === 0) {
    err('rawdata 中未找到 bs_order 文件');
    info('请将订单表（文件名以 bs_order 开头）放入: ' + RAWDATA_DIR);
    process.exit(1);
  }
  if (mdFiles.length === 0) {
    err('rawdata 中未找到 masterdata 文件');
    info('请将主数据表（文件名包含 masterdata）放入: ' + RAWDATA_DIR);
    process.exit(1);
  }

  ok('数据文件检查通过');
  info('订单表: ' + bsFiles[0]);
  info('主数据: ' + mdFiles[0]);

  const remotes = run('git remote -v', { silent: true, ignoreError: true });
  if (!remotes || !remotes.includes('github.com')) {
    warn('GitHub 远程仓库未配置');
    info('请先执行以下命令配置远程仓库：');
    console.log(C.yellow + '  git remote add origin https://github.com/你的用户名/fucaso-dashboard.git' + C.reset);
    console.log(C.yellow + '  git branch -M main' + C.reset);
    info('配置完成后重新运行本脚本');
    process.exit(1);
  }
  ok('GitHub 远程仓库已配置');

  // ─── 步骤 2：生成内置数据 ───
  step(2, '生成内置数据 (build-data.js)');

  const dataPath = path.join(ROOT, 'data-inline.js');
  let oldHash = '';
  if (fs.existsSync(dataPath)) {
    oldHash = run('git hash-object data-inline.js', { silent: true });
  }

  try {
    run('node build-data.js');
  } catch (e) {
    err('build-data.js 执行失败');
    console.error(e.message);
    process.exit(1);
  }

  if (!fs.existsSync(dataPath)) {
    err('data-inline.js 生成失败');
    process.exit(1);
  }

  // ─── 步骤 3：检查数据是否有变化 ───
  step(3, '检查数据变化');

  const newHash = run('git hash-object data-inline.js', { silent: true });
  if (oldHash && oldHash === newHash) {
    warn('data-inline.js 内容未发生变化');
    info('如果 Excel 已更新但数据没变，请检查文件内容');
    info('无需部署，脚本退出');
    process.exit(0);
  }

  ok(oldHash ? '检测到数据变化，准备部署' : '首次生成数据，准备部署');

  // ─── 步骤 4：Git 提交并推送 ───
  step(4, 'Git 提交并推送到 GitHub');

  const now = new Date();
  const dateStr = now.getFullYear() + '-' +
    String(now.getMonth() + 1).padStart(2, '0') + '-' +
    String(now.getDate()).padStart(2, '0');
  const timeStr = String(now.getHours()).padStart(2, '0') + ':' +
    String(now.getMinutes()).padStart(2, '0');
  const commitMsg = 'data: ' + dateStr + ' 销量日报更新';

  try {
    run('git add data-inline.js');
    run('git commit -m "' + commitMsg + '"');
    ok('Git 提交成功: ' + commitMsg);
  } catch (e) {
    warn('Git 提交可能已存在或无变化');
  }

  try {
    run('git push origin main');
    ok('已推送到 GitHub');
  } catch (e) {
    try {
      run('git push origin master');
      ok('已推送到 GitHub (master 分支)');
    } catch (e2) {
      err('推送到 GitHub 失败');
      console.error(e2.message);
      info('请检查网络连接或 GitHub 权限');
      process.exit(1);
    }
  }

  // ─── 步骤 5：完成输出 ───
  step(5, '完成');

  let pagesUrl = '';
  let repoUrl = '';
  try {
    const remoteUrl = run('git remote get-url origin', { silent: true });
    const match = remoteUrl.match(/github\.com[:\/]([^\/]+)\/([^\/\.]+)/);
    if (match) {
      pagesUrl = 'https://' + match[1] + '.github.io/' + match[2] + '/';
      repoUrl = 'https://github.com/' + match[1] + '/' + match[2];
    }
  } catch (e) {
    pagesUrl = 'https://你的用户名.github.io/fucaso-dashboard/';
    repoUrl = 'https://github.com/你的用户名/fucaso-dashboard';
  }

  log('\n' + C.bold + C.green + '🎉 部署完成！' + C.reset);
  log(C.gray + '────────────────────────────────────────' + C.reset);
  log('提交时间: ' + C.cyan + dateStr + ' ' + timeStr + C.reset);
  log('提交信息: ' + C.cyan + commitMsg + C.reset);
  log(C.gray + '────────────────────────────────────────' + C.reset);
  log('GitHub 仓库: ' + C.blue + repoUrl + C.reset);
  log('Dashboard 链接: ' + C.blue + pagesUrl + C.reset);
  log(C.gray + '────────────────────────────────────────' + C.reset);
  log('⏳ GitHub Pages 部署约需 1-2 分钟，请稍后刷新页面查看');
  log('');
}

main();
