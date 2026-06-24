# 采购单自动下载脚本 — CLAUDE.md

## 项目概述

通过 **Kimi WebBridge** 操作真实 Chrome 浏览器，自动登录盒马供应商平台，填写查询条件，逐页导出采购单 Excel 文件。无需 Selenium、无需关闭 Chrome、无需管理 ChromeDriver。

## 项目结构

```
browser_automation/
├── main.py                     # 主入口，完整流程编排（WebBridge 版）
├── requirements.txt            # 依赖：requests
├── README.md                   # 使用文档
├── config/
│   └── settings.py             # 所有可配置参数（日期、供应商、延迟等）
├── utils/
│   ├── webbridge_client.py     # WebBridge daemon HTTP API 封装
│   ├── helpers.py              # 通用工具：日志、日期计算、下载等待
│   └── driver_setup.py         # [已废弃] 旧 Selenium 驱动，保留参考
├── downloads/                  # （不再使用，下载走 Chrome 默认目录）
└── logs/                       # 运行日志，按时间戳命名
```

## 核心架构决策 & 原因

### 1. 选择器：按标签文字匹配，不按索引/类名

| 做法 | 示例 | 原因 |
|------|------|------|
| ❌ 类名匹配 | `button.next-btn-primary` | 页面上有 5 个同 class 按钮，querySelector 取第一个是"批量确认"而非"查询" |
| ❌ 数字索引 | `.next-row[1]`、`rows[4]` | 页面增减一行就全乱 |
| ✅ 标签文字 | `textContent.includes('创建日期')`、`textContent === '查询'` | 除非页面改文字，否则不会错 |

### 2. 分页 class 检测：用 `classList.contains` 不可以用 `cls.includes`

所有按钮的 class 都有 `next-btn`（Next UI 框架前缀）。`cls.includes('next')` 会误匹配所有按钮。必须用 `b.classList.contains('next')` 精确匹配独立 class。

### 3. 总页数检测：逐页翻到底

分页组件只显示滑动窗口（如 3 个页码按钮），不代表真实总数。通过点击"下一页"箭头直到 disabled，读出当前页码 = 真实总页数。

### 4. 下载目录：直接走 Chrome Downloads

WebBridge 通过真实 Chrome 操作，文件自动下载到 `%USERPROFILE%\Downloads`。不再通过 CDP 修改下载路径（CDP `Page.setDownloadBehavior` 在某些版本不稳定）。

### 5. 采购单状态：手动操作

`hippo-select-multiple` 组件的下拉选项渲染机制特殊，经过 DOM click、CDP 原生鼠标事件、父页面 overlay 搜索等多种尝试均无法自动操作。改为脚本暂停提示用户手动勾选，按回车继续。

### 6. 日期优先级：CLI > config > 自动

```
python main.py --start 2026-06-06 --end 2026-06-12   # CLI 优先
python main.py --start 2026-06-06 --add 6              # --add：start + N 天
```

不传则在 `settings.py` 中配置 `DELIVERY_DATE_START/END`，都不填则自动计算（今天-7 天 ~ 今天）。创建日期 = 要求到货 start - `CREATE_OFFSET_DAYS` 天。

### 7. 日期填写：分两次 evaluate + blur

Next UI DatePicker 在同一个 evaluate 中连续设两个值会导致 React 清掉第一个。必须分两次调用，每次 `focus() → click() → nativeSetter → input event → change event → blur()`。

### 8. 全选：点 checkbox 的父级 `<label>` 而非 `<input>`

Next UI 的 click 事件挂在 `<label>` 上。直接点 `<input>` 不会触发 React 的全选逻辑。

### 9. 弹窗关闭：只点 × 链接

对话框有关闭链接 `a.next-dialog-close` 和「去确认」按钮。只点 ×，不点任何文字按钮，避免误导航到退货确认页面。

### 10. 导出后关闭新 tab

导出按钮会 `window.open` 弹出一个退货确认新 tab。导出后检测 tab 数量变化，找到多余 tab 并关闭。

## 关键技术约束

- **Python 版本**: 3.10+
- **浏览器**: Chrome + Kimi WebBridge 扩展 + daemon 运行中
- **操作系统**: Windows
- **登录**: Chrome 中已有盒马登录态即可，无需脚本处理登录
- **iframe**: 直接导航到 iframe 内页 URL (`purchaseList.html`)，不再需要 switch 操作
- **页面组件**: 盒马当前使用 **Next UI (Fusion Design)**（`next-*` 前缀）+ **Hippo 组件**（`hippo-*` 前缀，如状态选择器）
- **WebBridge daemon**: `localhost:10086`，通过 `requests` 发 JSON POST

## 目标网站信息

- **门户首页**: `https://portalpro.hemaos.com/?storeTag=STORE_MANAGEMENT`
- **采购单列表（iframe 内页）**: `https://portalpro.hemaos.com/pages/supplierPlatformNew/purchaseList.html`

## 主流程（main.py）

1. 创建 WebBridgeClient，连接到 daemon (`localhost:10086`)
2. 直接导航到采购单列表 iframe URL（浏览器已登录）
3. 关闭页面弹窗（只点 ×）
4. 填写查询条件：
   - 供应商（输入关键词，选下拉匹配项）
   - 创建日期 / 要求到货日期（按标签文字定位行，分两次 evaluate 设值）
   - 采购单状态（暂停，用户手动勾选后按回车继续）
   - 导出 EXCEL 设置（勾选越库类型 checkbox）
5. 点击查询按钮（按文字"查询"精确匹配），等待表格加载
6. 逐页点"下一页"直到 disabled，读出真实总页数
7. 回到第一页，逐页循环：
   - 全选（点 `<label>` 包裹的 checkbox）
   - 导出 Excel（按文字"导出Excel"匹配按钮）
   - 等待文件出现在 `Downloads`
   - 翻页（优先点页码按钮，备用 next 箭头）
   - 关闭导出弹出的新 tab
8. 完成（不关闭浏览器）

## 常见修改点

| 需求 | 修改位置 |
|------|---------|
| 修改供应商关键词 | `config/settings.py` → `SUPPLIER_KEYWORD` / `SUPPLIER_NAME` |
| 修改日期范围 | 命令行 `--start --end` 或 `config/settings.py` → `DELIVERY_DATE_START/END` |
| 调整日期偏移 | `config/settings.py` → `CREATE_OFFSET_DAYS` |
| 修改采购单状态 | `config/settings.py` → `PURCHASE_STATUS_WANTED` |
| 调整操作等待时间 | `config/settings.py` 中的 `DELAY_*` 常量 |
| 修改 daemon 端口 | `config/settings.py` → `WEBBRIDGE_PORT` |
| 更新页面选择器 | `main.py` 中各 `evaluate()` 内的 JS 代码 |
| 修改翻页逻辑 | `main.py` → `go_to_next_page()` |
| 修改全选逻辑 | `main.py` → `export_current_page()` 中的 label click 部分 |

## 调试建议

- 运行日志保存在 `logs/` 目录
- 若选择器失效，按 F12 检查元素，页面使用 Next UI（`next-*`）和 Hippo（`hippo-*`）
- daemon 连接失败：`kimi-webbridge status` 检查，`kimi-webbridge start` 启动
- 扩展显示"浏览器助手未就绪"：检查注册表 `HKCU\Software\Google\Chrome\NativeMessagingHosts\com.kimi.webbridge`
- 下载超时可增大 `DELAY_DOWNLOAD`
- 页数不对：检查是否手动勾选了全部需要的采购单状态
- 可先用 WebBridge snapshot/evaluate 手动探索页面结构再改选择器

## 下一阶段目标

| 目标 | 优先级 | 说明 |
|------|--------|------|
| 状态下拉自动化 | 高 | 尝试通过父页面 + iframe 或 React fiber 方式自动勾选状态 |
| CDP 静默下载 | 中 | `Page.setDownloadBehavior` 设置下载不弹窗 |
| 第 3 页翻页修复 | 中 | 分页滑动窗口导致特定页码按钮缺失时的后备方案 |
| daemon 自动启动 | 低 | 内置到 `webbridge_client.py`，连接失败时自动拉起 daemon |
| 多浏览器支持 | 低 | Edge 的 NativeMessagingHosts 注册路径不同 |
| 命令行 `--status` | 低 | 支持命令行指定采购单状态，跳过手动步骤 |

## 不要做的事

- 不要在脚本运行时手动操作同一浏览器 tab，会干扰自动化
- 不要将账号密码硬编码进代码
- 旧 `driver_setup.py` 已废弃，新代码不再依赖 Selenium
- 不要用 `cls.includes('next')` 匹配 class，改用 `classList.contains('next')`
- 不要用数字索引定位 `.next-row`，改用标签文字匹配
