import os
import logging
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

logger = logging.getLogger(__name__)


def _resolve_chromedriver_service() -> Service:
    """按优先级获取 chromedriver：本地 conda 包 → webdriver-manager 下载"""
    # 1) 尝试 conda 安装的 chromedriver_binary
    try:
        from chromedriver_binary import chromedriver_filename
        chromedriver_path = chromedriver_filename
        if os.path.isfile(chromedriver_path):
            logger.info(f"使用本地 chromedriver: {chromedriver_path}")
            return Service(chromedriver_path)
    except ImportError:
        pass

    # 2) 尝试 webdriver-manager 下载
    try:
        path = ChromeDriverManager().install()
        if os.path.isfile(path):
            logger.info(f"使用 webdriver-manager 下载的 chromedriver: {path}")
            return Service(path)
    except Exception as e:
        logger.warning(f"webdriver-manager 下载失败: {e}")

    # 3) 回退到 PATH
    logger.info("回退到 PATH 中的 chromedriver")
    return Service()


def create_driver(user_data_dir: str, profile: str, download_dir: str) -> webdriver.Chrome:
    """
    创建有头 Chrome，使用真实用户 Profile（含已保存密码），
    并应用反检测参数隐藏自动化特征。
    """
    options = Options()

    # 显式指定 Chrome 路径
    options.binary_location = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

    # ── 反自动化检测 ──────────────────────────────────────────────────
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    # ── 模拟真实浏览器环境 ────────────────────────────────────────────
    # 注意：--disable-extensions 会与真实 Profile 冲突导致 Chrome 崩溃，已移除
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-infobars")
    options.add_argument("--start-maximized")

    # ── 下载设置：自动保存，不弹窗 ───────────────────────────────────
    os.makedirs(download_dir, exist_ok=True)
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        # 允许多文件同时下载（部分版本需要）
        "profile.default_content_setting_values.automatic_downloads": 1,
    }
    options.add_experimental_option("prefs", prefs)

    # ── 启动 Driver ──────────────────────────────────────────────────
    service = _resolve_chromedriver_service()
    driver = webdriver.Chrome(service=service, options=options)

    # ── 执行 JS 覆盖 navigator.webdriver ────────────────────────────
    driver.execute_cdp_cmd(
        "Page.addScriptToEvaluateOnNewDocument",
        {
            "source": """
                Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
                window.chrome = { runtime: {} };
                Object.defineProperty(navigator, 'plugins', {get: () => [1,2,3,4,5]});
                Object.defineProperty(navigator, 'languages', {get: () => ['zh-CN','zh','en']});
            """
        },
    )

    logger.info("Chrome 启动成功（反检测模式）")
    return driver
