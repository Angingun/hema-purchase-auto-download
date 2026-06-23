# 采购单自动下载脚本

## 📁 项目结构

```
browser_automation/
├── main.py                  # 主脚本（入口）
├── requirements.txt
├── config/
│   └── settings.py          # ⚙️ 所有配置参数（先改这里）
├── utils/
│   ├── driver_setup.py      # Chrome 启动 & 反检测
│   └── helpers.py           # 通用工具函数
├── downloads/               # 默认下载目录（可在 settings.py 修改）
└── logs/                    # 运行日志
```

---

## 🚀 快速开始

### 第一步：安装 Python 依赖

```bash
pip install -r requirements.txt
```

### 第二步：修改配置文件 `config/settings.py`

| 参数 | 说明 | 如何获取 |
|------|------|---------|
| `CHROME_USER_DATA_DIR` | Chrome 用户数据目录 | 浏览器地址栏输入 `chrome://version`，找「个人资料路径」，复制其**父目录** |
| `CHROME_PROFILE` | Profile 名称 | 同上，「个人资料路径」最后一段（如 `Default` 或 `Profile 1`） |
| `DOWNLOAD_DIR` | 文件保存位置 | 自定义绝对路径 |

**示例：**
```python
CHROME_USER_DATA_DIR = r"C:\Users\张三\AppData\Local\Google\Chrome\User Data"
CHROME_PROFILE = "Default"
DOWNLOAD_DIR = r"C:\Users\张三\Desktop\采购单导出"
```

### 第三步：运行脚本

```bash
python main.py
```

---

## ⚠️ 注意事项

1. **运行前必须关闭所有 Chrome 窗口**
   Chrome 不允许两个进程同时使用同一个 Profile，否则报错。

2. **Chrome 保存密码的工作原理**
   脚本加载你的真实 Chrome Profile，网站会自动填充已保存的密码并登录。
   如果网站需要手动操作（如验证码），脚本会暂停等你处理。

3. **日期参数**
   - 默认自动计算：结束日期=今天，开始日期=今天-7天
   - 手动指定：在 `main.py` 最后改为 `run(start_date="2024-06-01", end_date="2024-06-08")`

4. **文件命名规则**
   每页导出文件自动重命名为 `采购单_第001页_143022.xlsx`，防止覆盖。

5. **如果选择器失效**
   网站更新后 CSS 选择器可能变化，按 F12 重新检查元素，更新 `main.py` 中对应的选择器字符串。

---

## 🐛 常见问题

| 问题 | 解决方法 |
|------|---------|
| `DevToolsActivePort file doesn't exist` | 确保关闭了所有 Chrome 窗口后再运行 |
| 找不到供应商选项 | 检查 `SUPPLIER_KEYWORD` 和 `SUPPLIER_NAME` 是否与网站完全一致 |
| 下载超时 | 增大 `settings.py` 中的 `DELAY_DOWNLOAD` 值 |
| 日期填写失败 | 网站可能使用日期选择器而非直接输入，需调整 `_fill_date_input` 函数 |
