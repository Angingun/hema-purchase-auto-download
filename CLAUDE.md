# 采购单自动下载脚本 — CLAUDE.md

## 项目概述

通过 **Kimi WebBridge** 操作真实 Chrome 浏览器，自动登录盒马供应商平台，填写查询条件，逐页导出采购单 Excel 文件。无需 Selenium、无需关闭 Chrome、无需管理 ChromeDriver。

## 项目结构

```
browser_automation/
├── main.py                     # 主入口，完整流程编排（WebBridge 版）
├── requirements.txt            # 依赖：requests
├── config/
│   └── settings.py             # 所有可配置参数（daemon、URL、日期等）
├── utils/
│   ├── webbridge_client.py     # WebBridge daemon HTTP API 封装
│   ├── helpers.py              # 通用工具：日志、日期计算、下载等待
│   └── driver_setup.py         # [已废弃] 旧 Selenium 驱动，保留参考
├── downloads/                  # 下载文件输出目录
└── logs/                       # 运行日志，按时间戳命名
```

## 关键技术约束

- **Python 版本**: 3.10+
- **浏览器**: Chrome + Kimi WebBridge 扩展 + daemon 运行中
- **操作系统**: Windows
- **登录**: Chrome 中已有盒马登录态即可，无需脚本处理登录
- **iframe**: 直接导航到 iframe 内页 URL (`purchaseList.html`)，不再需要 switch 操作
- **页面组件**: 盒马当前使用 **Next UI (Fusion Design)**，类名前缀 `next-`

## 目标网站信息

- **门户首页**: `https://portalpro.hemaos.com/?storeTag=STORE_MANAGEMENT`
- **采购单列表（iframe 内页）**: `https://portalpro.hemaos.com/pages/supplierPlatformNew/purchaseList.html`

## 主流程（main.py）

1. 创建 WebBridgeClient，连接到 daemon (`localhost:10086`)
2. 直接导航到采购单列表 iframe URL（浏览器已登录）
3. 填写查询条件（均在 WebBridge evaluate/click/fill 中完成）：
   - 供应商（输入关键词，选下拉匹配项）
   - 创建日期（开始=今天-7天，结束=今天）
   - 要求到货日期（同上）
   - 采购单状态（多选审核通过/部分发货/发货完成等）
4. 点击查询，等待结果加载
5. 获取总页数
6. 逐页循环：全选 → 导出 Excel → 等待下载 → 重命名 → 翻页
7. 完成（不关闭浏览器）

## 常见修改点

| 需求 | 修改位置 |
|------|---------|
| 修改下载目录 / daemon 端口 | `config/settings.py` |
| 调整操作等待时间 | `config/settings.py` 中的 `DELAY_*` 常量 |
| 修改供应商关键词 | `config/settings.py` → `SUPPLIER_KEYWORD` / `SUPPLIER_NAME` |
| 手动指定日期 | `main.py` 最后一行 `run(start_date=..., end_date=...)` |
| 更新页面选择器 | `main.py` 中各 `evaluate()` 内的 JS 代码 |
| 修改翻页逻辑 | `main.py` → `go_to_next_page()` |

## 调试建议

- 运行日志保存在 `logs/` 目录
- 若选择器失效，按 F12 检查元素，页面使用 Next UI（`next-*` 前缀）
- daemon 连接失败：检查扩展是否已打开、daemon 是否运行
- 下载超时可增大 `DELAY_DOWNLOAD`，或检查浏览器下载设置
- 可先用 WebBridge snapshot/evaluate 手动探索页面结构再改选择器

## 不要做的事

- 不要在脚本运行时手动操作同一浏览器 tab，会干扰自动化
- 不要将账号密码硬编码进代码，密码由 Chrome 管理
- 旧 `driver_setup.py` 已废弃，新代码不再依赖 Selenium