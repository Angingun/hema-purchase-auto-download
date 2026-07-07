# Plan

## Goal

将盒马采购单自动下载脚本改造成“真实 Chrome 登录态 + 插件桥接 + 模拟人工操作 + 强校验”的稳定方案。

本轮测试日期固定为：

- 要求到货日期开始：`2026-07-04`
- 要求到货日期结束：`2026-07-06`

优先不走 CDP 主流程，不优先绕过浏览器接口。Codex Chrome 原生插件用于开发诊断；脚本运行时继续基于现有 WebBridge 类运行架构。

## Current Context

- 项目根目录：`C:\Users\Qingrun\Documents\01_Projects\01_Auto_download`
- 主入口：`browser_automation/main.py`
- 配置文件：`browser_automation/config/settings.py`
- WebBridge 封装：`browser_automation/utils/webbridge_client.py`
- 通用工具：`browser_automation/utils/helpers.py`
- 当前已知问题：
  - 翻页后可能未等当前页数据真正加载完成就导出。
  - 下载完成判断只看新文件出现，缺少文件大小稳定和 Excel 内容校验。
  - 当前页码、表格行数、导出文件之间没有建立可审计日志。
  - 采购单状态仍需人工选择。
- 运行约束：
  - Chrome 必须保持登录态。
  - 不把 CDP 或直接接口作为主路径。
  - 不回退到 Selenium / ChromeDriver。
  - 不覆盖用户已有未提交修改。

## Interface Contracts

### `PageState`

```python
PageState = dict[str, object]
```

字段约定：

- `page_num: int`：当前激活页码。
- `row_count: int`：当前页面可见表格数据行数。
- `loading: bool`：页面或表格是否仍处于加载状态。
- `first_row_key: str | None`：当前页首行稳定标识，用于判断翻页后数据是否变化。
- `raw: dict`：调试用原始 DOM 派生信息。

### `DownloadResult`

```python
DownloadResult = dict[str, object]
```

字段约定：

- `ok: bool`：下载是否成功。
- `path: str | None`：下载文件完整路径。
- `filename: str | None`：下载文件名。
- `size_bytes: int | None`：最终文件大小。
- `stable: bool`：文件大小是否已稳定。
- `error: str | None`：失败原因。

### `ExcelVerifyResult`

```python
ExcelVerifyResult = dict[str, object]
```

字段约定：

- `ok: bool`：Excel 是否可解析且通过基本校验。
- `path: str`：被校验文件路径。
- `sheet_names: list[str]`：工作表名称。
- `row_count: int | None`：解析出的数据行数。
- `column_count: int | None`：解析出的列数。
- `error: str | None`：失败原因。

### `get_page_state(wb: WebBridgeClient) -> PageState`

参数：

- `wb`：当前 WebBridge 浏览器会话。

返回：

- 当前页面状态，必须包含 `page_num`、`row_count`、`loading`、`first_row_key`。

约定：

- 不修改页面状态。
- 只读取 DOM。
- 若无法识别页码，`page_num` 返回 `0`，并在 `raw` 中记录可见分页文本。

### `wait_table_ready(wb: WebBridgeClient, expected_page: int, timeout: int = 30) -> PageState`

参数：

- `wb`：当前 WebBridge 浏览器会话。
- `expected_page`：期望激活的页码。
- `timeout`：最大等待秒数。

返回：

- 最终稳定的 `PageState`。

约定：

- 必须等待 loading 消失。
- 必须确认 `page_num == expected_page`。
- 必须等待 `row_count` 和 `first_row_key` 连续两次采样稳定。
- 超时必须抛出明确异常或返回带错误信息的状态，不能静默继续导出。

### `go_to_page(wb: WebBridgeClient, target_page: int, timeout: int = 30) -> PageState`

参数：

- `wb`：当前 WebBridge 浏览器会话。
- `target_page`：目标页码。
- `timeout`：最大等待秒数。

返回：

- 目标页加载稳定后的 `PageState`。

约定：

- 优先点击可见页码按钮。
- 若目标页码按钮不可见，使用下一页按钮逐页前进。
- 点击后必须调用 `wait_table_ready()`。
- 如果目标页无法到达，必须失败并记录当前页状态。

### `wait_download_complete(download_dir: str, before_files: set[str], timeout: int = 120) -> DownloadResult`

参数：

- `download_dir`：Chrome 下载目录。
- `before_files`：点击导出前的文件集合。
- `timeout`：最大等待秒数。

返回：

- `DownloadResult`。

约定：

- 忽略 `.crdownload` 临时文件。
- 必须等待新文件出现。
- 必须等待文件大小连续两次稳定。
- 文件大小为 0 时失败。

### `verify_excel_file(path: str) -> ExcelVerifyResult`

参数：

- `path`：下载后的 Excel 文件路径。

返回：

- `ExcelVerifyResult`。

约定：

- 文件必须存在。
- 文件必须能被解析。
- 至少包含 1 个 sheet。
- 行数和列数必须可记录。
- 若缺少 Excel 解析依赖，任务必须明确更新依赖文件。

## Task Checklist

- [ ] 单元 1：记录当前仓库状态与执行边界
  - 类型：主任务
  - 输入：当前 git 工作区、现有项目文件
  - 输出：执行前状态记录，明确哪些文件已有未提交修改
  - 涉及文件：无代码修改；只读 `git status --short`
  - 接口契约：无
  - 验收标准：Claude Code 能说明执行前有哪些已修改或未跟踪文件，并承诺不回滚用户修改。
- [ ] 单元 2：新增或更新项目协作说明
  - 类型：可并行子任务
  - 输入：当前 `CLAUDE.md`、`browser_automation/README.md`、本计划
  - 输出：根目录 `AGENTS.md`
  - 涉及文件：`AGENTS.md`
  - 接口契约：无
  - 验收标准：`AGENTS.md` 明确说明优先保留真实 Chrome 登录态、WebBridge 运行架构、Codex Chrome 仅用于开发诊断、不要优先走 CDP/接口/Selenium。
- [ ] 单元 3：固定本轮测试日期参数
  - 类型：主任务
  - 输入：测试日期 `2026-07-04` 到 `2026-07-06`
  - 输出：脚本可通过 CLI 使用该日期运行，不强制改写默认配置
  - 涉及文件：`browser_automation/main.py`、`browser_automation/config/settings.py`
  - 接口契约：现有 `run(start_date: str | None, end_date: str | None, add_days: int = 0)` 保持兼容
  - 验收标准：可使用 `python main.py --start 2026-07-04 --end 2026-07-06` 运行；不破坏 config 默认日期逻辑。
- [ ] 单元 4：实现页面状态读取
  - 类型：主任务
  - 输入：当前 WebBridge 会话、采购单列表页面 DOM
  - 输出：`get_page_state(wb)` 函数
  - 涉及文件：`browser_automation/main.py` 或新增 `browser_automation/utils/page_state.py`
  - 接口契约：`get_page_state(wb: WebBridgeClient) -> PageState`
  - 验收标准：函数不点击页面，只读取状态；日志能输出当前页码、可见行数、loading 状态、首行标识。
- [ ] 单元 5：实现表格稳定等待
  - 类型：主任务
  - 输入：`get_page_state(wb)`、期望页码
  - 输出：`wait_table_ready(wb, expected_page, timeout=30)` 函数
  - 涉及文件：`browser_automation/main.py` 或 `browser_automation/utils/page_state.py`
  - 接口契约：`wait_table_ready(wb: WebBridgeClient, expected_page: int, timeout: int = 30) -> PageState`
  - 验收标准：只有页码正确、loading 消失、行数和首行标识稳定后才返回；超时必须终止当前页导出并记录错误。
- [ ] 单元 6：重构翻页逻辑
  - 类型：主任务
  - 输入：当前页码、目标页码、分页组件 DOM
  - 输出：`go_to_page(wb, target_page, timeout=30)` 函数
  - 涉及文件：`browser_automation/main.py`
  - 接口契约：`go_to_page(wb: WebBridgeClient, target_page: int, timeout: int = 30) -> PageState`
  - 验收标准：每次翻页后必须确认当前激活页码等于目标页；无法到达目标页时失败，不继续导出错误页。
- [ ] 单元 7：增强当前页导出前校验
  - 类型：主任务
  - 输入：目标页码、`PageState`
  - 输出：导出前审计日志
  - 涉及文件：`browser_automation/main.py`
  - 接口契约：复用 `PageState`
  - 验收标准：每页导出前日志至少包含目标页码、实际页码、可见行数、首行标识；实际页码不匹配时不点击导出。
- [ ] 单元 8：实现下载完成强校验
  - 类型：主任务
  - 输入：下载目录、导出前文件集合
  - 输出：`wait_download_complete(download_dir, before_files, timeout=120)`
  - 涉及文件：`browser_automation/utils/helpers.py`
  - 接口契约：`wait_download_complete(download_dir: str, before_files: set[str], timeout: int = 120) -> DownloadResult`
  - 验收标准：必须等待 `.crdownload` 消失和文件大小稳定；返回结构化结果；下载失败时记录明确原因。
- [ ] 单元 9：实现 Excel 文件校验
  - 类型：可并行子任务
  - 输入：下载后的 Excel 文件路径
  - 输出：`verify_excel_file(path)` 函数
  - 涉及文件：`browser_automation/utils/helpers.py`、`browser_automation/requirements.txt`
  - 接口契约：`verify_excel_file(path: str) -> ExcelVerifyResult`
  - 验收标准：可解析 `.xlsx` 文件并返回 sheet 名、行数、列数；如需依赖，`requirements.txt` 必须增加明确依赖，例如 `openpyxl`。
- [ ] 单元 10：将下载与 Excel 校验接入导出流程
  - 类型：主任务
  - 输入：`export_current_page()`、`wait_download_complete()`、`verify_excel_file()`
  - 输出：每页结构化导出结果
  - 涉及文件：`browser_automation/main.py`
  - 接口契约：`export_current_page(wb: WebBridgeClient, page_num: int, download_dir: str) -> dict[str, object] | None`
  - 验收标准：每页导出后日志包含文件名、大小、Excel 行数、校验状态；校验失败时不中断后续页前必须清晰记录失败页码。
- [ ] 单元 11：生成运行汇总
  - 类型：主任务
  - 输入：每页导出结果列表
  - 输出：最终汇总日志
  - 涉及文件：`browser_automation/main.py`
  - 接口契约：导出结果字段至少包含 `page_num`、`ui_row_count`、`file_path`、`file_size`、`excel_row_count`、`ok`、`error`
  - 验收标准：运行结束时输出总页数、成功页数、失败页数、UI 可见行数合计、Excel 行数合计、失败页详情。
- [ ] 单元 12：使用 Codex Chrome 做非主流程诊断
  - 类型：可并行子任务
  - 输入：用户已登录的 Chrome、测试日期 `2026-07-04` 到 `2026-07-06`
  - 输出：诊断记录，说明页面 loading、分页、下载按钮、状态多选框的实际行为
  - 涉及文件：可写入运行日志或单独诊断记录；不要求改代码
  - 接口契约：无
  - 验收标准：诊断结果能支持或修正 `get_page_state()`、`wait_table_ready()`、`go_to_page()` 的 DOM 判断条件。
- [ ] 单元 13：执行 1 次小范围真实运行验证
  - 类型：主任务
  - 输入：CLI 参数 `--start 2026-07-04 --end 2026-07-06`
  - 输出：下载文件、运行日志、汇总结果
  - 涉及文件：`browser_automation/logs/`、Chrome 下载目录
  - 接口契约：无新增
  - 验收标准：脚本能完成运行；每一页都有页码确认、行数记录、下载文件、Excel 校验结果；如果失败，日志足够定位失败环节。
- [ ] 单元 14：更新 README 的运行与排障说明
  - 类型：可并行子任务
  - 输入：最终实现行为、测试命令
  - 输出：更新后的 `browser_automation/README.md`
  - 涉及文件：`browser_automation/README.md`
  - 接口契约：无
  - 验收标准：README 明确说明测试命令、Chrome 登录态要求、人工选择采购单状态、下载与 Excel 校验日志的位置。

## Files Not To Touch

- 不修改账号、密码、Cookie、Token 或任何登录凭据。
- 不删除或回滚用户已有改动。
- 不修改 `.git/`。
- 不修改 Chrome 用户 Profile 文件。
- 不启用 Selenium / ChromeDriver 作为主流程。
- 不把 CDP 或直接后端接口调用改成主流程。
- 不自动关闭用户非本任务相关 Chrome 标签页。

## Acceptance Criteria

- 可以使用 `python main.py --start 2026-07-04 --end 2026-07-06` 执行测试。
- 脚本运行时继续依赖真实 Chrome 登录态和现有 WebBridge 类方案。
- 每页导出前必须确认当前页码与目标页码一致。
- 每页导出前必须确认表格加载完成且状态稳定。
- 每个下载文件必须经过文件大小稳定校验。
- 每个下载文件必须经过 Excel 可解析校验。
- 最终日志必须能判断是否漏页、重复页、下载失败或 Excel 内容异常。
- Codex Chrome 插件只作为开发诊断工具，不作为 Python 脚本运行时依赖。

## Test Plan

- 执行前确认 Chrome 已登录盒马供应商平台。
- 使用日期范围 `2026-07-04` 到 `2026-07-06`。
- 运行命令：`python main.py --start 2026-07-04 --end 2026-07-06`。
- 人工完成采购单状态选择后继续。
- 检查日志：
  - 查询条件是否正确。
  - 总页数是否识别。
  - 每页目标页码和实际页码是否一致。
  - 每页可见行数是否大于 0，空页除外。
  - 每页下载文件是否存在且大小稳定。
  - 每页 Excel 是否可解析。
- 检查汇总：
  - 成功页数 + 失败页数等于计划处理页数。
  - Excel 行数合计合理。
  - 失败页有明确错误原因。

## Risks / Assumptions

- 阻塞问题：当前无法确认盒马页面在 `2026-07-04` 到 `2026-07-06` 是否一定有数据；如果无数据，测试只能验证空结果路径。
- 阻塞问题：采购单状态多选仍需人工选择；如果人工选择不完整，导出数据范围会受影响。
- 假设 Chrome 已保持有效登录态。
- 假设 Kimi WebBridge daemon 可正常连接真实 Chrome。
- 假设下载文件为 `.xlsx` 或可被 Excel 解析库读取的格式。
- 假设页面分页组件仍基于 Next UI。
- 风险：网站可能忽略 synthetic click；若发生，后续再评估坐标点击或 WebBridge 扩展能力，而不是直接改 CDP 主流程。
- 风险：Chrome 下载目录中可能有同名文件；实现必须通过导出前后文件集合和文件稳定性判断新文件。
