# FUCASO 销量日报 Dashboard · 部署与使用指南

> 福可苏 FUCASO 销量日报 · 网页版 Dashboard（暗黑科技风）

## 项目结构

```
fucaso-dashboard-web/
├── index.html          # 主页面（Dashboard UI + 登录层）
├── dashboard.js        # 核心逻辑（双模式：内置数据 / 上传模式）
├── build-data.js       # Node.js 脚本：Excel → 内置数据
├── deploy.js           # ⭐ 一键部署脚本：数据更新 + GitHub 推送
├── data-inline.js      # 内置数据文件（由 build-data.js 自动生成）
├── rawdata/            # 放置 Excel 数据文件（不上传 GitHub）
├── _headers            # 安全响应头配置
├── .gitignore
└── README.md           # 本文档
```

> ⭐ `data-inline.js` 是**核心数据文件**。网站展示的数据全部来自此文件。  
> ⭐ `deploy.js` 是**一键部署脚本**，运行后自动完成「生成数据 → Git 提交 → 推送到 GitHub」全流程。

---

## 两种使用模式

| 模式 | 适用场景 | 说明 |
|------|---------|------|
| **预置数据模式** ⭐ 推荐 | 日常运营 | 管理员本地更新 Excel → 运行 `deploy.js` → 团队成员打开链接直接看结果 |
| **临时上传模式** | 紧急调试 / 无 Node 环境 | 在浏览器中拖拽上传 Excel，立即生成报表（数据不保存，刷新需重新上传） |

---

## 🚀 管理员日常更新（只需一步）

将最新的 Excel 放入 `rawdata/` 后，运行：

```bash
node deploy.js
```

脚本自动完成：
1. ✅ 检查 `rawdata/` 中的 bs_order + masterdata 文件
2. ✅ 运行 `build-data.js` 生成 `data-inline.js`
3. ✅ 检测数据是否有变化（避免无意义提交）
4. ✅ 自动 `git add → commit → push`
5. ✅ 输出 GitHub Pages 链接

---

## 首次使用：配置指南（只需一次）

### 1. 准备数据文件

在 `rawdata/` 放入两个 Excel：

```
rawdata/
├── bs_order_xxxx.xlsx    # 订单表（文件名以 bs_order 开头）
└── masterdata.xlsx       # 主数据表（固定文件名）
```

### 2. 配置 Git 并推送代码

```bash
# 进入项目目录
cd fucaso-dashboard-web

# 初始化仓库
git init

# 添加远程仓库（替换为你的 GitHub 用户名）
git remote add origin https://github.com/你的用户名/fucaso-dashboard.git

# 设置分支名
git branch -M main

# 首次提交所有文件
git add index.html dashboard.js build-data.js deploy.js _headers .gitignore rawdata/.gitkeep
git commit -m "init: fucaso dashboard"

# 推送
git push -u origin main
```

### 3. 开启 GitHub Pages

1. 打开 [github.com](https://github.com)，进入你的 `fucaso-dashboard` 仓库
2. 点击 **Settings** → 左侧 **Pages**
3. Source：Branch 选 `main`，Folder 选 `/(root)` → 点击 **Save**
4. 等待 1-2 分钟，页面上方显示绿色链接：
   ```
   ✅ Your site is live at https://你的用户名.github.io/fucaso-dashboard/
   ```

完成以上配置后，以后每次更新数据只需运行 `node deploy.js`。

---

## 团队成员使用

1. 打开 Dashboard 链接
2. 输入密码 → 点击「进入看板」
3. **直接看到最新销量日报**，无需上传任何文件
4. 可点击「📸 下载日报 PNG」保存图片

---

## 修改密码

打开 `dashboard.js`，找到第 19 行：

```javascript
const ACCESS_PASSWORD = 'fucaso2026';  // ⚠️ 修改此处为你要设置的密码
```

修改后保存，推送到 GitHub 即可生效。

---

## 无 Node.js 环境的替代方案

如果暂时没有 Node.js 环境，可用浏览器「临时上传模式」生成 `data-inline.js`：

1. 用浏览器打开 `index.html`（本地双击打开即可）
2. 登录后显示上传页面 → 拖拽上传 bs_order + masterdata
3. 点击「生成销量日报」→ Dashboard 正常显示
4. **关键步骤**：按 F12 打开控制台，输入：
   ```javascript
   copy('window.BUILTIN_DATA = ' + JSON.stringify(window.BUILTIN_DATA, null, 2) + ';')
   ```
5. 粘贴到剪贴板的内容就是 `data-inline.js` 的完整内容
6. 新建 `data-inline.js`，粘贴保存，然后上传到 GitHub

---

## 绑定自定义域名（可选）

### GitHub 端
1. Settings → Pages → **Custom domain**
2. 输入：`dashboard.fucaso.com`（或其他含 `fucaso` 的域名）
3. 勾选 **Enforce HTTPS**

### DNS 端

| 记录类型 | 主机记录 | 记录值 |
|---------|---------|--------|
| CNAME | dashboard | 你的用户名.github.io |

---

## 常见问题 FAQ

### Q1: 运行 deploy.js 提示 "GitHub 远程仓库未配置"？
A: 首次使用需要配置 Git remote，执行：
```bash
git remote add origin https://github.com/你的用户名/fucaso-dashboard.git
```

### Q2: 运行 deploy.js 提示 "内容未发生变化"？
A: 说明 rawdata/ 中的 Excel 文件和上次生成的数据相同。如果确认已更新 Excel，请检查文件内容是否正确保存。

### Q3: 网站打开显示上传页面，而不是 Dashboard？
A: `data-inline.js` 缺失或路径错误。请确保该文件已推送到 GitHub，并等待 1-2 分钟缓存刷新。

### Q4: 更新数据后网站没有变化？
A: GitHub Pages 有缓存延迟。尝试强制刷新 `Ctrl + F5` 或链接后加 `?v=2`。

### Q5: 可以设置多个密码吗？
A: 当前只支持单密码。如需多用户管理，建议升级到 Cloudflare Access。

### Q6: 前端密码安全吗？
A: 适合小团队内部使用（"防君子不防小人"）。密码存在于 JS 源码中。如需更安全方案，请使用 Cloudflare Access + 邮箱白名单。

---

## 进阶：升级到 Cloudflare Access

如需企业级访问控制（邮箱白名单、SSO、访问日志）：

1. 将代码部署到 Cloudflare Pages
2. Cloudflare Zero Trust → Access → 创建应用
3. 配置邮箱白名单或企业 SSO
4. 删除 `dashboard.js` 中的 `initLogin` 函数

> Cloudflare Access 免费版支持 50 个用户，零月费。
