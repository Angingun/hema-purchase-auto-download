"""二分排查 Selenium flag 导致 Chrome Profile 崩溃"""
import sys, os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from chromedriver_binary import chromedriver_filename

PROFILE_DIR = r"C:\Users\Qingrun\AppData\Local\Google\Chrome\User Data"
PROFILE_NAME = "Default"


def test(options: Options, label: str) -> bool:
    """返回 True 表示启动成功"""
    try:
        driver = webdriver.Chrome(service=Service(chromedriver_filename), options=options)
        driver.quit()
        print(f"  [OK] {label}")
        return True
    except Exception as e:
        msg = str(e)[:80]
        print(f"  [FAIL] {label}: {msg}")
        return False


def build_base() -> Options:
    opts = Options()
    opts.binary_location = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    opts.add_argument(f"--user-data-dir={PROFILE_DIR}")
    opts.add_argument(f"--profile-directory={PROFILE_NAME}")
    return opts


# --- 测试序列 ---
print("\n===== 逐一排查 Selenium Flag =====\n")

# 0: 纯 Profile，无任何附加 flag
print("0. 仅 Profile（控制组）:")
opts = build_base()
test(opts, "仅 Profile")

# 1: --no-sandbox
print("\n1. + --no-sandbox:")
opts = build_base()
opts.add_argument("--no-sandbox")
test(opts, "--no-sandbox")

# 2: --disable-blink-features
print("\n2. + --disable-blink-features=AutomationControlled:")
opts = build_base()
opts.add_argument("--disable-blink-features=AutomationControlled")
test(opts, "disable-blink-features")

# 3: excludeSwitches + useAutomationExtension
print("\n3. + excludeSwitches & useAutomationExtension:")
opts = build_base()
opts.add_experimental_option("excludeSwitches", ["enable-automation"])
opts.add_experimental_option("useAutomationExtension", False)
test(opts, "excludeSwitches+useAutomationExtension")

# 4: --disable-dev-shm-usage
print("\n4. + --disable-dev-shm-usage:")
opts = build_base()
opts.add_argument("--disable-dev-shm-usage")
test(opts, "--disable-dev-shm-usage")

# 5: download prefs
print("\n5. + download prefs:")
opts = build_base()
os.makedirs(r"C:\Users\Qingrun\Desktop\采购单导出", exist_ok=True)
opts.add_experimental_option("prefs", {
    "download.default_directory": r"C:\Users\Qingrun\Desktop\采购单导出",
    "download.prompt_for_download": False,
})
test(opts, "download prefs")

print("\n===== 排查完成 =====")
