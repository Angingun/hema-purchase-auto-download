# AGENTS.md -- 盒马采购单自动下载项目

> 给未来 AI Agent 的行为指导和架构说明。

## 1. Agent 行为规则

### 保留真实 Chrome 登录态
- 所有浏览器操作必须通过 Kimi WebBridge（真实 Chrome 扩展桥接）进行。
- 不做任何形式的自动登录。Chrome 的盒马供应商平台登录态由用户在运行前手动完成。
- 不要通过代码注入 Cookie、Token 或凭证。

### WebBridge 优先架构
- 主流程始终使用 `webbridge_client.py` 封装的 HTTP API（`localhost:10086`）。
- 所有页面交互（点击、输入、导航、JS 执行）通过 `WebBridgeClient.evaluate()` 或 `fill()` / `click()` 方法。
- 不要将 CDP（Chrome DevTools Protocol）作为主操作路径。`cdp()` 方法仅在调试或诊断时使用。

### Codex Chrome / MCP Chrome-DevTools 仅限开发诊断
- 本项目的 MCP chrome-devtools 工具（`take_snapshot`、`evaluate_script`、`list_network_requests` 等）仅供 Agent 在开发阶段探索页面 DOM、调试选择器、分析网络请求时使用。
- 不要将 MCP Chrome 工具作为 Python 脚本运行时的依赖。运行时只依赖 `webbridge_client.py` + requests。

### 不走 Selenium / ChromeDriver / 后端 API
- 项目已从 Selenium + ChromeDriver 迁移到 WebBridge 方案。旧的 `driver_setup.py` 保留参考但已废弃。
- 不要尝试绕过浏览器直接调用盒马后端 API。
- 不要重新引入 Selenium 依赖。

### 不动用户凭证和文件
- 不修改、不读取、不提交任何包含账号密码、Cookie、Token 或登录凭据的文件。
- 不删除或回滚用户已有的未提交修改（git status 中的 modified/untracked 文件）。
- 不修改 Chrome 用户 Profile 文件。
- 不自动关闭用户非本任务相关的 Chrome 标签页。

---

## 2. 关键架构决策

### 元素匹配：按标签文字，不按索引/类名

| 做法 | 示例 | 原因 |
|------|------|------|
| 文字匹配 | `textContent.includes('创建日期')` | 页面增减行或修改样式不会失效 |
| 文字匹配按钮 | `textContent === '查询'` | 避免 match 到同 class 的其他按钮 |

不要用数字索引定位 `.next-row[i]`。

### Class 检测：用 `classList.contains()` 不是 `includes()`

```javascript
// Correct
b.classList.contains('next')

// Wrong -- matches every Next UI button
b.className.includes('next')
```

### 分页：逐页点"下一页"直到 disabled
分页组件只显示滑动窗口，不代表真实总数。循环点击 next 箭头直到按钮 disabled，记录总点击数。

### DatePicker：分两次 evaluate + blur 设值
Next UI DatePicker 不能在一个 evaluate 中连续设两个值（React 会清掉第一个）。必须每次 `focus() -> click() -> native value setter -> input event -> change event -> blur()`。

### 全选：点 `<label>` 不是 `<input>`
Next UI 的 click 事件挂在 `<label>` 上。点 `<input>` 不会触发 React 的全选逻辑：
```javascript
const label = cb.closest('label');
if (label) label.click();
```

### 弹窗关闭：只点 × 链接
对话框的关闭链接是 `a.next-dialog-close`。只点这个，不点任何文字按钮（如「去确认」），避免误导航到退货确认页面。

### 导出后关闭新 Tab
导出按钮会 `window.open` 弹出新 tab。导出后检测 tab 数量变化，找到退货确认 tab 并关闭，然后切回采购单列表 tab。

---

## 3. 接口契约

以下为 `PLAN.md` 定义的接口规范，实现时必须遵守。

### PageState
```python
PageState = dict[str, object]
# page_num: int        -- 当前激活页码（无法识别时返回 0）
# row_count: int       -- 可见表格数据行数
# loading: bool        -- 页面/表格是否仍在加载
# first_row_key: str | None -- 首行稳定标识
# raw: dict            -- 调试用原始 DOM 派生信息
```

### DownloadResult
```python
DownloadResult = dict[str, object]
# ok: bool             -- 下载是否成功
# path: str | None     -- 下载文件完整路径
# filename: str | None -- 文件名
# size_bytes: int | None -- 最终文件大小
# stable: bool         -- 文件大小是否已稳定
# error: str | None    -- 失败原因
```

### ExcelVerifyResult
```python
ExcelVerifyResult = dict[str, object]
# ok: bool             -- Excel 是否可解析且通过基本校验
# path: str            -- 被校验文件路径
# sheet_names: list[str] -- 工作表名称
# row_count: int | None   -- 解析出的数据行数
# column_count: int | None -- 解析出的列数
# error: str | None    -- 失败原因
```

### `get_page_state(wb: WebBridgeClient) -> PageState`
- 纯读取，不修改 DOM。
- 必须返回 page_num、row_count、loading、first_row_key。
- 无法识别页码时返回 0。

### `wait_table_ready(wb, expected_page, timeout=30) -> PageState`
- 等待 loading 消失。
- 确认 page_num == expected_page。
- 等待 row_count 和 first_row_key 连续两次采样稳定。
- 超时抛出 TimeoutError。

### `go_to_page(wb, target_page, timeout=30) -> PageState`
- 优先点击可见页码按钮。不可见时逐页前进。
- 翻页后调用 `wait_table_ready()`。
- 无法到达目标页时抛出异常。

### `wait_download_complete(download_dir, before_files, timeout=120) -> DownloadResult`
- 忽略 `.crdownload` 临时文件。
- 等待新文件出现。
- 等待文件大小连续两次采样稳定。
- 文件大小为 0 时失败。

### `verify_excel_file(path) -> ExcelVerifyResult`
- 文件必须存在且可被 openpyxl 解析。
- 至少包含 1 个 sheet。
- 记录行数、列数、sheet 名。

### `export_current_page(wb, page_num, download_dir, page_state) -> dict`
- 导出前校验：确认页码匹配、行数 > 0。
- 点击全选 -> 点击导出 -> 强校验下载。
- 关闭导出弹出的新 tab。
- 返回结构化结果，包含 page_num、ok、ui_row_count、file_path、file_size、error 等字段。

---

## 4. 相关文件

| 文件 | 说明 |
|------|------|
| `browser_automation/main.py` | 主入口，完整流程编排 |
| `browser_automation/config/settings.py` | 所有可配置参数（供应商、日期、延迟等） |
| `browser_automation/utils/webbridge_client.py` | WebBridge daemon HTTP API 封装 |
| `browser_automation/utils/helpers.py` | 日志、日期计算、下载等待、Excel 校验 |
| `browser_automation/utils/page_state.py` | 页面状态读取和表格稳定等待 |
| `browser_automation/utils/driver_setup.py` | [已废弃] 旧 Selenium 驱动，仅参考 |
| `browser_automation/logs/` | 运行日志输出目录 |
| `PLAN.md` | 详细开发计划和接口契约定义 |
| `AGENTS.md` | 本文件 -- 给 AI Agent 的行为指南 |

---

## 5. 重要约束

- **Python**: 3.10+
- **依赖**: requests（HTTP 调用）、openpyxl（Excel 校验）
- **操作系统**: Windows（路径硬编码 %USERPROFILE%）
- **WebBridge daemon**: `localhost:10086`，通过 `kimi-webbridge start` 自动拉起
- **Chrome 要求**: 必须已登录盒马供应商平台（`portalpro.hemaos.com`）
- **下载路径**: Chrome 默认下载目录（`%USERPROFILE%\Downloads`），不通过 CDP 修改
- **页面框架**: Next UI (Fusion Design) + Hippo 组件
- **操作延迟**: `config/settings.py` 中 DELAY_SHORT/MEDIUM/LONG/DOWNLOAD 可调

### 运行方式
```bash
cd browser_automation
python main.py                                    # 使用 settings.py 日期
python main.py --start 2026-07-04 --end 2026-07-06  # CLI 参数优先
python main.py --start 2026-07-04 --add 6           # start + N 天
```

### 采购单状态多选（已知阻塞项）
`hippo-select-multiple` 组件的下拉选项渲染机制特殊，DOM click、CDP 原生事件、React fiber 操作均不可靠。当前方案：脚本暂停，提示用户在浏览器中手动勾选后按回车继续。

### 已知问题
- 翻页后滚动窗口可能导致特定页码按钮不可见（第 3 页翻页问题）。
- 下载目录同名文件冲突（通过 before/after 文件集合差分判断新文件）。
- 页面可能无数据，空页应跳过导出而非报错。
