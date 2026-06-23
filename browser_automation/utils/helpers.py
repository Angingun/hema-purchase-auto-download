import logging
import time
import os
from datetime import datetime, timedelta
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

logger = logging.getLogger(__name__)


def setup_logging(log_dir: str):
    """初始化日志（同时输出到控制台和文件）"""
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, f"run_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler(log_file, encoding="utf-8"),
        ],
    )
    logger.info(f"日志文件: {log_file}")


def get_date_range(end_date_str: str = None, days_back: int = 7):
    """
    返回 (start_date, end_date) 字符串，格式 YYYY-MM-DD
    end_date 默认今天，start_date = end_date - days_back 天
    """
    if end_date_str:
        end = datetime.strptime(end_date_str, "%Y-%m-%d")
    else:
        end = datetime.today()
    start = end - timedelta(days=days_back)
    return start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")


def wait_and_find(driver, css_selector: str, timeout: int = 15):
    """等待元素可见并返回"""
    return WebDriverWait(driver, timeout).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, css_selector))
    )


def wait_clickable(driver, css_selector: str, timeout: int = 15):
    """等待元素可点击并返回"""
    return WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, css_selector))
    )


def safe_click(driver, css_selector: str, timeout: int = 15, use_js: bool = False):
    """安全点击：等待可点击后点击，失败时用 JS 备用"""
    el = wait_clickable(driver, css_selector, timeout)
    try:
        if use_js:
            driver.execute_script("arguments[0].click();", el)
        else:
            ActionChains(driver).move_to_element(el).pause(0.3).click().perform()
    except Exception:
        driver.execute_script("arguments[0].click();", el)
    return el


def type_text(driver, css_selector: str, text: str, clear_first: bool = True, timeout: int = 15):
    """模拟人工输入文字"""
    el = wait_and_find(driver, css_selector, timeout)
    el.click()
    if clear_first:
        el.send_keys(Keys.CONTROL + "a")
        el.send_keys(Keys.DELETE)
        time.sleep(0.2)
    for char in str(text):
        el.send_keys(char)
        time.sleep(0.05)  # 模拟人工逐字输入
    return el


def switch_to_iframe(driver, css_selector: str, timeout: int = 20):
    """切换到指定 iframe（先切回主框架）"""
    driver.switch_to.default_content()
    iframe = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, css_selector))
    )
    driver.switch_to.frame(iframe)
    logger.info(f"已切换到 iframe: {css_selector}")


def wait_for_new_file(download_dir: str, before_files: set, timeout: int = 60) -> str | None:
    """等待下载目录出现新文件，返回文件路径"""
    deadline = time.time() + timeout
    while time.time() < deadline:
        current = set(os.listdir(download_dir))
        new_files = current - before_files
        # 过滤掉 .crdownload 临时文件
        completed = [f for f in new_files if not f.endswith(".crdownload")]
        if completed:
            path = os.path.join(download_dir, completed[0])
            logger.info(f"下载完成: {path}")
            return path
        time.sleep(1)
    logger.warning("等待下载超时")
    return None
