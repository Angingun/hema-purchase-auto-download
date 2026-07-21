# AGENTS.md -- 盒马采购单自动下载项目

> 给未来 AI Agent 的项目记忆、行为边界和架构说明。当前项目使用真实 Chrome 登录态 + Kimi WebBridge 完成盒马供应商平台采购单逐页导出。

## 1. 当前进展摘要

### 已新增 / 强化的模块与功能

- `browser_automation/utils/page_state.py`
  - 新增页面状态读取模块。
  - 提供 `get_page_state(wb)` 和 `wait_table_ready(wb, expected_page, timeout=30, previous_order_hash=None)`。
  - 只读取主可见表格 body，规避 Next UI 固定列/隐藏表格导致的重复行误判。
  - 提取当前页采购单号 `order_ids`，并生成订单集合 hash `order_hash`。
  - 判断当前页码、可见行数、loading 状态、首行 key、是否可点击下一页。

- `browser_automation/utils/helpers.py`
  - 强化下载完成检测：等待新文件出现，忽略 `.crdownload`，等待文件大小稳定。
  - 强化 Excel 校验：除 sheet/行列数外，解析采购单号集合并生成 `order_hash`。
  - Excel 校验结果可与页面 `order_hash` 对比，判断“下载文件是否就是当前页”。

- `browser_automation/main.py`
  - 主流程改为“边导出边翻页”，不再先全量探测总页数。
  - 每页导出前等待页面稳定，导出后校验 Excel 中订单集合与页面订单集合一致。
  - 翻页只点击 Next 箭头，并等待页码 + 订单 hash 同时变化。
  - 增加重复页 hash 防护，防止翻页失败后重复导出上一页。
  - 增强 WebBridge 启动流程：执行 `kimi-webbridge.exe start` 后调用 `wb.wait_ready()`，确认 daemon 与 Chrome 扩展真正握手成功。
  - 新增 `python main.py --check-webbridge`：依次检查可执行文件、daemon 端口和扩展握手，不导航页面、不下载。
  - 新增采购单状态自动选择：逐项输入筛选并点击精确菜单项，最终反读已选标签校验；失败时回退人工调整。
  - Chrome 保持默认下载目录；每页 Excel 校验通过后移动到 `DOWNLOAD_DIR`，目标重名时追加序号且不覆盖历史文件。

- `browser_automation/utils/webbridge_client.py`
  - 新增 `wait_ready(timeout=30, interval=1.0)`。
  - 构造 `WebBridgeClient` 不代表连接成功；必须通过 `list_tabs()` 等实际命令确认 daemon/扩展就绪。
  - 重启电脑后 Kimi WebBridge 扩展常见“未就绪”问题，会在业务流程前提前失败并给出明确提示。

- 文档
  - 新增/更新 `PLAN.md`：Claude Code 可执行的实施合同。
  - 新增/更新 `EXECUTION.md`：执行记录与验证情况。
  - 更新 `README.md`：运行方式、测试命令、WebBridge 未就绪排障、校验说明。
  - 更新本文件 `AGENTS.md`：给后续 Agent 的项目规则与下一阶段方向。

### 已验证结果

- 测试日期：`2026-07-04` 到 `2026-07-06`。
- 新流程成功导出 4 页，没有再出现旧流程的第 1/3 页重复和缺页问题。
- 状态自动选择已在真实页面验证：五项缩减为前三项后校验通过，再自动补回两项并校验五项完全一致。
- 自动导出文件 `(54)-(57)` 与手工文件 `(49)-(52)` 对比：前 3 页完全一致；第 4 页差异确认来自筛选时间/状态不一致，不再作为下载完整性问题处理。
- 语法检查通过：

```bash
python -m py_compile browser_automation/main.py browser_automation/utils/page_state.py browser_automation/utils/helpers.py browser_automation/utils/webbridge_client.py
```

## 2. 重要架构决策和原因

### 使用真实 Chrome 登录态，不自动登录

- 盒马供应商平台依赖用户登录态，项目不保存、不注入、不读取账号密码、Cookie、Token。
- 运行前用户需要保证 Chrome 已登录 `portalpro.hemaos.com`。
- 不修改 Chrome Profile，不自动关闭用户非本任务标签页。

### 运行时继续使用 Kimi WebBridge，不把 Codex Chrome 插件写进脚本

- Python 脚本运行时使用 `WebBridgeClient` 调用本地 daemon：`http://127.0.0.1:10086/command`。
- Codex Chrome 原生插件只用于开发期观察 DOM、截图、验证网页结构；不能作为 `main.py` 的运行时依赖。
- 原因：脚本需要可被用户独立运行，而 Codex Chrome 插件是 Agent 开发环境能力，不是项目依赖。

### 不优先走 CDP / Selenium / 后端接口

- 项目已从 Selenium + ChromeDriver 迁移到 WebBridge。
- 不重新引入 Selenium；`utils/driver_setup.py` 仅保留历史参考。
- 不绕过浏览器直接调用盒马后端 API。
- CDP 仅作为诊断或极端交互 fallback，不作为主路径。

### 分页策略：边导出边翻页，不预扫总页数

旧问题来自先翻页探测总页数再回第一页，页面状态容易变旧或加载不完全，导致 Excel 下载重复/缺页。

当前策略：

1. 查询后等待第 1 页稳定。
2. 导出当前页并校验页面订单 hash == Excel 订单 hash。
3. 如果 Next 可用，点击下一页。
4. 等待页码递增且订单 hash 变化。
5. 重复直到 Next disabled 或达到 `MAX_PAGES`。

### 页面稳定判断：页码、loading、行数、首行 key、订单 hash 一起看

只看 DOM 行数不够，Next UI 会同时存在多个 `.next-table-body`，并且加载状态可能滞后。

当前 `wait_table_ready()` 要求：

- 当前页码等于预期页码。
- 表格 loading 结束。
- 主表可见行数 > 0。
- 行数、首行 key、订单 hash 连续采样稳定。
- 翻页场景下当前页订单 hash 必须不同于上一页。

### Excel 校验：校验订单集合，不只校验文件可打开

旧的“文件能打开、有行数”不足以证明下载的是当前页。

当前校验：

- Excel 可被 `openpyxl` 打开。
- 至少有一个 sheet。
- 记录行数、列数。
- 从首个 sheet 提取唯一采购单号集合。
- 计算订单集合 hash，并与页面 hash 对比。

### WebBridge 启动：自动 start + 显式 ready check

- 只执行 `kimi-webbridge.exe start` 不代表扩展已经连接。
- 重启电脑后扩展面板可能显示“未就绪”。
- 当前 `main.py` 会：
  - 调用 `kimi-webbridge.exe start`。
  - 构造 `WebBridgeClient`。
  - 调用 `wb.wait_ready()` 轮询 `list_tabs()`。
  - 如果仍未就绪，在导航前失败并提示打开 Chrome/扩展面板确认。
- 不自动执行 `stop` / `restart` / `uninstall`，避免杀掉正在运行的 daemon 或影响用户浏览器状态。

## 3. 当前项目结构

```text
01_Auto_download/
├── AGENTS.md                         # 本文件：Agent 项目记忆与行为规则
├── PLAN.md                           # Claude Code 实施合同与 checklist
├── EXECUTION.md                      # 本轮执行记录与验证状态
├── CLAUDE.md                         # 早期 Claude 项目说明，仍可参考但以 AGENTS.md 为准
└── browser_automation/
    ├── main.py                       # 主入口：日期计算、WebBridge 启动、查询、翻页、导出、汇总
    ├── requirements.txt              # Python 依赖
    ├── README.md                     # 用户运行说明与排障
    ├── cdp_download.py               # 历史 CDP 尝试，仅参考，不作为主流程
    ├── config/
    │   └── settings.py               # 供应商、日期、状态、延迟、最大页数等配置
    ├── utils/
    │   ├── webbridge_client.py       # Kimi WebBridge HTTP API 封装 + ready check
    │   ├── page_state.py             # 页面状态读取、表格稳定等待、订单 hash
    │   ├── helpers.py                # 日志、日期、下载等待、Excel 校验
    │   └── driver_setup.py           # 废弃 Selenium 参考代码
    └── logs/                         # 运行日志输出目录
```

## 4. 关键接口契约

### `PageState`

```python
PageState = dict[str, object]
# page_num: int              -- 当前激活页码，无法识别时为 0
# row_count: int             -- 主可见表格 body 的业务行数
# loading: bool              -- 页面或表格是否仍在加载
# first_row_key: str | None  -- 首行业务内容稳定标识
# order_ids: list[str]       -- 当前页唯一采购单号，保留页面顺序
# order_hash: str | None     -- 当前页采购单号集合 hash
# can_next: bool             -- 下一页按钮是否可用
# raw: dict                  -- 调试用 DOM 派生信息
```

### `get_page_state(wb: WebBridgeClient) -> PageState`

- 纯读取，不修改页面状态。
- 必须只统计主可见 `.next-table-body`。
- 必须返回页码、行数、loading、首行 key、订单集合、订单 hash、Next 可用性。

### `wait_table_ready(wb, expected_page, timeout=30, previous_order_hash=None) -> PageState`

- 等待指定页码稳定。
- 翻页时如果提供 `previous_order_hash`，必须确认新页订单 hash 已变化。
- 超时抛 `TimeoutError`，不能静默继续导出。

### `verify_excel_file(path) -> ExcelVerifyResult`

```python
ExcelVerifyResult = dict[str, object]
# ok: bool
# path: str
# sheet_names: list[str]
# row_count: int | None
# column_count: int | None
# order_ids: list[str]
# order_hash: str | None
# error: str | None
```

### `export_current_page(wb, page_num, download_dir, page_state=None) -> dict`

- 导出前校验当前页码与预期页码一致。
- 当前页行数为 0 时跳过导出。
- 全选当前页，点击导出 Excel。
- 等待下载完成并校验 Excel。
- Excel `order_hash` 必须等于页面 `order_hash` 才算成功。
- 返回结构化结果，包含页面行数、页面订单数、Excel 行数、Excel 订单数、文件路径、错误信息等。

### `WebBridgeClient.wait_ready(timeout=30, interval=1.0) -> None`

- 用于确认 daemon 与 Chrome 扩展真实可用。
- 成功时返回 `None`。
- 超时抛 `WebBridgeError`，错误信息应提示 Chrome 扩展未连接/未就绪。

### `select_purchase_statuses(wb, wanted, timeout=10, fallback_manual=True) -> dict`

- 自动移除目标集合之外的状态，并逐项输入筛选、点击缺失状态。
- 以第一个 `.hippo-select-multiple` 的已选标签作为最终事实来源。
- 返回 `ok`、`wanted`、`selected`、`missing`、`extra`、`options_seen`、`method`、`error`。
- 只有 `missing == []` 且 `extra == []` 时 `ok=True`；自动失败可回退人工调整并再次反读。

### `check_webbridge(timeout=10.0) -> bool`

- 执行一次幂等 daemon `start`，但不执行 `stop` / `restart`。
- 先验证 `WEBBRIDGE_PORT` 是否监听，再用 `wait_ready()` / `list_tabs()` 验证扩展握手。
- 成功返回 `True`；任一层失败返回 `False`，CLI 映射为退出码 `1`。
- 不导航页面、不读取业务 DOM、不下载文件。

## 5. 运行方式

```bash
cd browser_automation
python main.py
python main.py --start 2026-07-04 --end 2026-07-06
python main.py --start 2026-07-04 --add 6
```

运行前确认：

- Chrome 已打开。
- Chrome 已登录盒马供应商平台。
- Kimi WebBridge 扩展已安装并启用。
- 如果扩展显示“未就绪”，先运行：

```powershell
& "$env:USERPROFILE\.kimi-webbridge\bin\kimi-webbridge.exe" start
```

## 6. 已知问题与当前限制

### 采购单状态自动选择的限制

主流程会自动将采购单状态调整为 `PURCHASE_STATUS_WANTED`，并反读标签确认目标集合完全一致。若页面组件结构变化、选项未出现或点击后校验不一致，会暂停并回退人工调整；人工完成后仍必须通过反读校验，不会只相信回车确认。

### Kimi WebBridge 重启后未就绪

已经加入 `wait_ready()` 提前发现问题，但如果 Chrome 扩展/native messaging 本身没有连上，脚本无法强行修复。需要用户打开 Chrome 和扩展面板确认。

### `settings.py` 可能出现无文本 diff 的 modified 状态

之前观察到 `browser_automation/config/settings.py` 在 `git status` 中显示 modified，但 `git diff` 无文本内容。处理提交时不要无脑纳入，先确认是否真有内容变化。

## 7. 下一阶段目标

### 目标 1：查询条件快照

每次点击查询前/后记录：

- 创建日期范围。
- 要求到货日期范围。
- 采购单状态已选值。
- 供应商已选值。
- 页面总条数/分页文本（如果可读）。

原因：当手工导出与自动导出结果不一致时，可以快速判断是筛选条件差异还是下载/翻页问题。

### 目标 2：清理旧路径和文档一致性

- `get_total_pages()`、`go_to_page()` 已非主流程，可评估是否保留为诊断工具或删除。
- `cdp_download.py` 和 `driver_setup.py` 应继续标记为历史参考，避免后续 Agent 误用。
- `CLAUDE.md` 中旧结论应逐步同步到 `AGENTS.md` 或注明过期。

## 8. 给后续 Agent 的注意事项

- 改代码前先读 `PLAN.md`、`EXECUTION.md`、`AGENTS.md` 和当前 `git diff`。
- 不要回滚用户未提交修改。
- 不要把 Codex Chrome 插件能力写成项目运行时依赖。
- 不要为了“更稳定”重新引入 Selenium、ChromeDriver 或直接后端 API。
- 页面交互优先通过 `WebBridgeClient.evaluate()`，必要时通过 Codex Chrome 插件做开发期 DOM 诊断。
- 每次修改后至少运行：

```bash
python -m py_compile browser_automation/main.py browser_automation/utils/page_state.py browser_automation/utils/helpers.py browser_automation/utils/webbridge_client.py
```