# 采购单自动下载脚本

通过 Kimi WebBridge 操作真实 Chrome 浏览器，自动登录盒马供应商平台，填写查询条件，逐页导出采购单 Excel。

> **不再需要 Selenium、ChromeDriver、关闭 Chrome。**

## 前置条件

1. **Chrome** + 已登录盒马供应商平台 (`portalpro.hemaos.com`)
2. **Kimi WebBridge 扩展** 已安装并启用
3. **WebBridge daemon** 在运行（启动 Chrome 后打开扩展面板即自动拉起）

## 安装

```bash
pip install requests
```

## 配置

编辑 `config/settings.py`：

```python
# 供应商
SUPPLIER_KEYWORD = "282265890"
SUPPLIER_NAME = "KA佳农食品(上海)有限公司（新）"

# 要求到货日期（留空则自动使用"今天-7天 ~ 今天"）
DELIVERY_DATE_START = ""     # 如 "2026-06-06"
DELIVERY_DATE_END   = ""     # 如 "2026-06-12"

# 创建日期 = 要求到货 start 往前推 N 天
CREATE_OFFSET_DAYS = 7

# 采购单状态（脚本运行时会暂停让你手动勾选）
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
```

## 运行流程

1. 自动打开采购单列表页面（利用浏览器已有登录态，无需登录）
2. 关闭弹窗 → 填写供应商 → 填写日期 → 勾选导出设置
3. **暂停**：提示你在浏览器中手动勾选采购单状态，完成后按回车
4. 点击查询 → 检测总页数 → 逐页全选导出 Excel
5. 文件保存到 `C:\Users\<用户名>\Downloads`（Chrome 默认下载目录）

## 项目结构

```
browser_automation/
├── main.py                     # 主入口
├── requirements.txt            # 依赖：requests
├── config/
│   └── settings.py             # 所有可配置参数
├── utils/
│   ├── webbridge_client.py     # WebBridge daemon HTTP API 封装
│   ├── helpers.py              # 日志、日期计算、下载等待
│   └── driver_setup.py         # [已废弃] 旧 Selenium 驱动
├── logs/                       # 运行日志
└── README.md
```

## 常见问题

| 问题 | 解决方法 |
|------|---------|
| 无法连接到 WebBridge daemon | 打开 Chrome 中 WebBridge 扩展面板，daemon 会自动启动 |
| 页数不对 | 检查是否手动勾选了全部需要的采购单状态 |
| 下载超时 | 增大 `config/settings.py` 中 `DELAY_DOWNLOAD` |
| 选择器失效 | 页面更新了 Next UI 组件，按 F12 检查元素更新选择器 |
