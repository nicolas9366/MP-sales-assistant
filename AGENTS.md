# AGENTS.md · bonyu.de 官方主页开发与 Netcup 自动部署指南

> 本指南供 Codex 及所有协作 Agent 在维护与编辑 `http://bonyu.de/` 主页内容时遵守。

---

## 一、项目与域名映射

- **线上主页域名**：`http://bonyu.de/`（或 `https://bonyu.de/`）
- **本地代码仓库**：`/Users/dengyu/github/MP-sales-assistant`
- **GitHub 远端仓库**：`git@github.com:nicolas9366/MP-sales-assistant.git`（默认主分支：`main`）
- **自动部署机制**：本仓库通过 GitHub Webhook 绑定了 Netcup Web 服务器。当代码推送到 `origin main` 后，Netcup 服务器将在 3~5 秒内自动拉取最新代码并实时生效。

---

## 二、核心主页文件职责与结构

本仓库根目录即为 `http://bonyu.de/` 的线上站点根目录，核心前端主页文件职责如下：

| 文件 | 职责说明 | 技术栈 / 依赖 |
| :--- | :--- | :--- |
| `index.html` | 网站主页入口文件（主看板页面结构） | Tailwind CSS (CDN), FontAwesome, Chart.js |
| `app.js` | 前端核心交互逻辑、图表渲染与状态驱动 | 原生 JavaScript (ES6+), Chart.js |
| `styles.css` | 页面自定义样式表与微调样式 | CSS3 (含暗色系看板与工业精密风样式) |
| `sample_data.js` | 客户档案、销售数据与看板前端展示数据 | JavaScript Object / JSON |
| `config.json` | 站点与数据源配置信息 | JSON |

> **注意**：
> - `ingest_agent.log`、`data_loader.py`、`ingest_agent.py` 为数据清洗与后台采集脚本，非主页前端页面文件。编辑主页时切勿将无关临时日志提交入库。
> - `niko-broadcast/` 播客门户统一由备用站点 `https://kyrmax.de/niko-broadcast/` 提供，本仓库专用于 `bonyu.de` 主页看板与销售业务前端。

---

## 三、编辑主页内容与 Netcup 自动同步 SOP

Codex 在编辑 `bonyu.de` 主页时，请遵循以下标准三步流程：

### 1. 检查工作区状态并编辑
在修改前先确认当前工作区状态：
```bash
git -C /Users/dengyu/github/MP-sales-assistant status
```
根据需求编辑 `index.html`、`app.js`、`styles.css`、`sample_data.js` 或 `config.json`。

### 2. 仅暂存目标文件并执行 Git 三连 Push 触发自动部署
严禁使用 `git add .`（避免误提交 `ingest_agent.log` 等临时日志）。明确指定修改的文件：
```bash
# 示例：仅暂存主页相关文件
git -C /Users/dengyu/github/MP-sales-assistant add index.html app.js styles.css

# 规范化 Commit 信息
git -C /Users/dengyu/github/MP-sales-assistant commit -m "[codex] feat: update bonyu.de homepage content"

# 推送到 GitHub 触发 Netcup 自动拉取
git -C /Users/dengyu/github/MP-sales-assistant push origin main
```

### 3. 线上验活（等待 3~5 秒后验证 HTTP 200）
推送完成后稍候数秒，通过 `curl` 检查线上站点响应：
```bash
curl -I -L -s --max-time 8 http://bonyu.de/ | head -n 15
```
确保返回 `HTTP/1.1 200 OK` 且站点内容如预期更新。
