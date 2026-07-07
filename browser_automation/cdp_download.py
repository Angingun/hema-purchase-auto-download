"""
CDP 翻页下载脚本
前置：Chrome 已通过 --remote-debugging-port=9222 启动，页面已查询出结果
功能：连接已打开的 Chrome → 找到采购单列表 tab → 逐页全选导出
"""
import os
import sys
import re
import time
import logging
from datetime import datetime

logger = logging.getLogger(__name__)


def run_download_loop():
    """假设页面已查询完成显示表格，只做逐页下载。"""
    import subprocess
    subprocess.run([sys.executable, "-m", "pip", "install", "playwright", "-q"], capture_output=True)

    from playwright.sync_api import sync_playwright

    dl_dir = os.path.expandvars(r"%USERPROFILE%\Downloads")

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp("http://127.0.0.1:9222")
        if not browser.contexts:
            logger.error("No browser contexts found")
            return

        # 找到采购单列表页
        page = None
        for ctx in browser.contexts:
            for pg in ctx.pages:
                if "purchaseList" in pg.url:
                    page = pg
                    break
            if page:
                break

        if not page:
            logger.error("未找到采购单列表页，请确认页面已打开")
            return

        logger.info("已连接: %s", page.url)

        # ── 1. 检测总页数 ──────────────────────────────────────────
        total_pages = 1
        for _ in range(500):
            can_next = page.evaluate("""() => {
                const pager = document.querySelector('.next-pagination');
                if (!pager) return false;
                const btns = pager.querySelectorAll('button');
                for (const b of btns) {
                    if (b.classList.contains('next'))
                        return !b.disabled && !b.classList.contains('disabled');
                }
                return false;
            }""")
            if not can_next:
                break
            page.evaluate("""() => {
                const pager = document.querySelector('.next-pagination');
                pager.querySelectorAll('button').forEach(b => {
                    if (b.classList.contains('next')) b.click();
                });
            }""")
            time.sleep(1)
            total_pages += 1

        # 读当前页码（应该是最后一页）
        current_text = page.evaluate("""() => {
            const pager = document.querySelector('.next-pagination');
            if (!pager) return '1';
            for (const b of pager.querySelectorAll('button')) {
                if (b.className.includes('current')) return b.textContent.trim();
            }
            return '?';
        }""")
        if current_text.isdigit():
            total_pages = int(current_text)

        logger.info("检测到 %d 页", total_pages)

        # 回到第 1 页
        page.evaluate("""() => {
            const pager = document.querySelector('.next-pagination');
            if (!pager) return;
            pager.querySelectorAll('button').forEach(b => {
                if (b.textContent.trim() === '1') b.click();
            });
        }""")
        time.sleep(2)

        # ── 2. 逐页导出 ──────────────────────────────────────────────
        for pg_num in range(1, total_pages + 1):
            logger.info("━━━ 第 %d/%d 页 ━━━", pg_num, total_pages)

            # 等待表格数据加载
            for _ in range(30):
                has_data = page.evaluate("""() => {
                    const loading = document.querySelector('.next-table-loading, .next-loading');
                    if (loading && loading.offsetParent !== null) return false;
                    return document.querySelectorAll('.next-table-body td').length > 0;
                }""")
                if has_data:
                    break
                time.sleep(0.5)
            else:
                logger.warning("第 %d 页无数据，跳过", pg_num)
                if pg_num < total_pages:
                    _go_next(page, pg_num)
                continue

            # 全选
            page.evaluate("""() => {
                const cbs = document.querySelectorAll('.next-table-header input[type="checkbox"]');
                for (const cb of cbs) {
                    const label = cb.closest('label');
                    if (label) {
                        if (cb.checked) label.click();
                        label.click();
                        return;
                    }
                }
            }""")
            time.sleep(1)

            # 点击导出
            before = set(os.listdir(dl_dir))
            page.evaluate("""() => {
                const btns = document.querySelectorAll('button');
                for (const b of btns) {
                    if (b.textContent.includes('导出Excel')) { b.click(); return; }
                }
            }""")

            # 等待下载
            new_file = None
            for _ in range(120):
                current = set(os.listdir(dl_dir))
                diff = current - before
                completed = [f for f in diff if not f.endswith('.crdownload')]
                if completed:
                    new_file = os.path.join(dl_dir, completed[0])
                    break
                time.sleep(1)

            if new_file:
                logger.info("✔ %s", new_file)
            else:
                logger.error("✗ 下载超时")

            # 翻到下一页
            if pg_num < total_pages:
                _go_next(page, pg_num)

        logger.info("完成！共 %d 页", total_pages)


def _go_next(page, current: int):
    """翻到下一页，等待数据加载。"""
    target = current + 1
    for attempt in range(3):
        time.sleep(1 + attempt * 2)
        clicked = page.evaluate(f"""() => {{
            const btns = document.querySelectorAll('.next-pagination button');
            for (const b of btns) {{
                if (b.textContent.trim() === '{target}' && !b.disabled) {{
                    b.click(); return true;
                }}
            }}
            return false;
        }}""")
        if clicked:
            time.sleep(3)
            return True
    return False


def setup_logging():
    log_dir = os.path.join(os.path.dirname(__file__), "logs")
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, f"cdp_run_{datetime.now():%Y%m%d_%H%M%S}.log")
    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[logging.StreamHandler(), logging.FileHandler(log_file, encoding="utf-8")],
    )
    logger.info("日志: %s", log_file)


if __name__ == "__main__":
    setup_logging()
    run_download_loop()
