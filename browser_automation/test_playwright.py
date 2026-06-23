"""测试 Playwright 能否用 Chrome Profile 启动"""
import sys
print("Testing Playwright with Chrome Profile...", flush=True)

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print("请先安装: pip install playwright && playwright install chromium", flush=True)
    sys.exit(1)

with sync_playwright() as p:
    try:
        context = p.chromium.launch_persistent_context(
            user_data_dir=r"C:\Users\Qingrun\AppData\Local\Google\Chrome\User Data",
            channel="chrome",
            headless=False,
            args=[
                "--profile-directory=Default",
            ],
        )
        page = context.pages[0] if context.pages else context.new_page()
        print("SUCCESS: Chrome launched with Profile!", flush=True)

        # 在已有 tab 中导航
        try:
            page.goto("https://portalpro.hemaos.com/?storeTag=STORE_MANAGEMENT", timeout=15000)
            print(f"导航成功: {page.url}", flush=True)
        except Exception as e:
            print(f"导航失败: {e}", flush=True)

        input("检查后按回车关闭...")
        context.close()
    except Exception as e:
        print(f"FAIL: {e}", flush=True)
