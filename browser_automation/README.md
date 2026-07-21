# 采购单自动下载脚本

通过 Kimi WebBridge 操作真实 Chrome 浏览器，自动登录盒马供应商平台，填写查询条件，逐页导出采购单 Excel。

> **不再需要 Selenium、ChromeDriver、关闭 Chrome。**

## 前置条件

1. **Chrome** + 已登录盒马供应商平台 (`portalpro.hemaos.com`)
2. **Kimi WebBridge 扩展** 已安装并启用
3. **WebBridge daemon** 可由脚本自动 `start`；重启电脑后如扩展显示未就绪，先打开 Chrome 并点击扩展面板确认连接

## 安装

```bash
pip install requests openpyxl
```

## 配置

编辑 `config/settings.py`：

```python
# WebBridge 本地端口（本机 10086 被 Windows 排除，因此改用 18086）
WEBBRIDGE_PORT = 18086

# 校验通过后的归档目录
DOWNLOAD_DIR = r"C:\Users\Qingrun\OneDrive\01_工作\10_盒马-佳农\05_Sales\03_OrdersbyStores_RawData"

# 供应商
SUPPLIER_KEYWORD = "282265890"
SUPPLIER_NAME = "KA佳农食品(上海)有限公司（新）"

# 要求到货日期（留空则自动使用"今天-7天 ~ 今天"）
DELIVERY_DATE_START = ""     # 如 "2026-06-06"
DELIVERY_DATE_END   = ""     # 如 "2026-06-12"

# 创建日期 = 要求到货 start 往前推 N 天
CREATE_OFFSET_DAYS = 7

# 采购单状态（脚本自动选择并反读校验）
PURCHASE_STATUS_WANTED = [
    "审核通过", "部分发货", "发货完成",
    "全部入库", "部分入库",
]
```

## 运行

```bash
# 使用 config/settings.py 中的日期
python main.py

# 命令行传参（优先级高于 config）
python main.py --start 2026-06-06 --end 2026-06-12

# --start + --add：结束日期 = 开始日期 + N 天
python main.py --start 2026-06-06 --add 6

# 也可以配合 --end 精确指定
python main.py --start 2026-06-06 --end 2026-06-12

# 仅检查 WebBridge，不打开网页或下载
python main.py --check-webbridge
```

## 测试运行

推荐使用的测试命令（仅导出指定日期范围，快速验证整体流程）：

```bash
python main.py --start 2026-07-04 --end 2026-07-06
```

## 运行流程

1. 自动打开采购单列表页面（利用浏览器已有登录态，无需登录）
2. 关闭弹窗 → 填写供应商 → 填写日期 → 自动选择并校验采购单状态 → 勾选导出设置
3. 自动选择失败时暂停，待人工调整后再次反读校验
4. 点击查询 → 读取结果区“共 N 条数据” → 逐页全选导出 Excel
5. Chrome 先下载到 `CHROME_DOWNLOADS_DIR`，校验通过后移动到 `DOWNLOAD_DIR`

## 校验与验证

脚本在每个环节都加入了自动校验，确保导出文件正确：

- **页面状态校验 (Page state check)** — 导出前验证当前页码与预期一致。日志记录 `Pre-export check`，包含目标页码、实际页码、行数及首行关键信息。
- **下载稳定性检查 (Download stability)** — 等待文件大小稳定后才视为下载完成，日志记录文件字节数。
- **Excel 内容校验 (Excel validation)** — 使用 `openpyxl` 解析已下载文件，验证其是否为有效 Excel。结果包含 Sheet 名称、行数和列数。
- **查询总量核验 (Query total validation)** — 查询后读取右上角“共 N 条数据”；全部导出后按本轮 Excel 中的唯一采购单号（`HPO`）汇总比对。数目不一致时，本次运行标记为失败。
- **运行汇总 (Run summary)** — 脚本结束时输出结构化汇总，包含总页数、成功/失败数量、UI 表行数合计、Excel 验证行数合计。

所有校验明细记录在 `logs/` 目录下的日志文件中。

## 项目结构

```
browser_automation/
├── main.py                     # 主入口，完整流程编排
├── requirements.txt            # 依赖：requests, openpyxl
├── config/
│   └── settings.py             # 所有可配置参数
├── utils/
│   ├── webbridge_client.py     # WebBridge daemon HTTP API 封装
│   ├── page_state.py           # 页面状态读取、表格稳定性等待
│   ├── helpers.py              # 日志、日期计算、下载等待、Excel 校验
│   └── driver_setup.py         # [已废弃] 旧 Selenium 驱动
├── logs/                       # 运行日志（含校验明细和汇总）
└── README.md
```

## 常见问题

| 问题 | 解决方法 |
|------|---------|
| 无法连接到 WebBridge daemon | 脚本会自动执行 `kimi-webbridge.exe start`；若仍失败，打开 Chrome 并点击 Kimi WebBridge 扩展面板确认连接 |
| 扩展显示“未就绪” | 先确保 Chrome 已打开，再运行 `& "$env:USERPROFILE\.kimi-webbridge\bin\kimi-webbridge.exe" start`；如果仍未就绪，重新打开扩展面板或按 Kimi WebBridge 官方说明修复 native messaging |
| WebBridge 健康检查 | 运行 `python main.py --check-webbridge`，按输出区分 daemon 未监听与扩展未握手 |
| 自定义端口后扩展未连接 | 连续点击扩展弹窗顶部 Kimi 图标 5 次，在“高级设置”把 Daemon WebSocket 地址改为 `ws://127.0.0.1:18086/ws`，测试后保存 |
| 状态自动选择失败 | 按日志提示手动调整；脚本会反读已选标签，完全匹配配置后才继续 |
| 下载超时 | 增大 `config/settings.py` 中 `DELAY_DOWNLOAD` |
| 选择器失效 | 页面更新了 Next UI 组件，按 F12 检查元素更新选择器 |
| Excel 校验失败 | 确认文件是否完整下载；检查 `openpyxl` 是否正确安装 |