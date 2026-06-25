"""
采购单自动下载脚本 (WebBridge 版)
网站: portalpro.hemaos.com
依赖: Kimi WebBridge daemon (localhost:10086) + Chrome 扩展
功能: 通过 WebBridge 操作真实浏览器，自动填写查询条件并逐页导出 Excel

前置条件:
  1. Chrome 已安装 Kimi WebBridge 扩展
  2. daemon 已启动（打开扩展面板即自动启动）
  3. Chrome 已登录盒马供应商平台
"""

import os
import subprocess
import sys
import time
import json
import logging
from datetime import datetime, timedelta

sys.path.insert(0, os.path.dirname(__file__))

from config.settings import (
    WEBBRIDGE_PORT, PURCHASE_LIST_URL,
    DOWNLOAD_DIR, CHROME_DOWNLOADS_DIR,
    SUPPLIER_KEYWORD, SUPPLIER_NAME,
    DELIVERY_DATE_START, DELIVERY_DATE_END, CREATE_OFFSET_DAYS,
    PURCHASE_STATUS_WANTED,
    DELAY_SHORT, DELAY_MEDIUM, DELAY_LONG, DELAY_DOWNLOAD, MAX_PAGES,
)
from utils.helpers import (
    setup_logging, get_date_range, wait_for_new_file,
)
from utils.webbridge_client import WebBridgeClient, WebBridgeError

logger = logging.getLogger(__name__)

SESSION = "hema-" + datetime.now().strftime("%H%M%S")


def _js_str(s: str) -> str:
    """将 Python 字符串安全地嵌入 JS 单引号字符串。"""
    return json.dumps(s, ensure_ascii=False)


def fill_search_form(wb: WebBridgeClient, create_start: str, create_end: str,
                     delivery_start: str, delivery_end: str):
    """填写查询条件：供应商、日期、状态、导出设置。"""

    # ── 1. 供应商 ────────────────────────────────────────────────────
    logger.info("  填写供应商: %s", SUPPLIER_KEYWORD)
    wb.fill('[placeholder*="供应商"]', SUPPLIER_KEYWORD)
    time.sleep(DELAY_MEDIUM)

    matched = wb.evaluate(f"""
        (() => {{
            const name = {_js_str(SUPPLIER_NAME)};
            const items = document.querySelectorAll(
                '.next-menu-item, .next-select-menu-item, ' +
                'li[role="option"], .next-comboBox-menu-item'
            );
            for (const item of items) {{
                if (item.textContent.includes(name)) {{
                    item.click();
                    return 'matched: ' + item.textContent.trim().substring(0, 40);
                }}
            }}
            if (items.length > 0) {{
                items[0].click();
                return 'fallback: ' + items[0].textContent.trim().substring(0, 40);
            }}
            return 'no items found';
        }})()
    """)
    logger.info(f"  供应商选择结果: {matched}")
    time.sleep(DELAY_SHORT)

    # ── 2. 创建日期 = 到货日期 start - offset ──────────────────────────
    logger.info("  填写创建日期: %s ~ %s", create_start, create_end)
    _fill_date_row(wb, "创建日期", create_start, create_end)

    # ── 3. 要求到货日期（从 config 读取）─────────────────────────────
    logger.info("  填写要求到货日期: %s ~ %s", delivery_start, delivery_end)
    _fill_date_row(wb, "要求到货", delivery_start, delivery_end)

    # ── 4. 采购单状态（多选）── 需手动操作 ──────────────────────────
    if PURCHASE_STATUS_WANTED:
        logger.info("  ⏳ 请在浏览器中手动选择采购单状态：")
        logger.info("     需要勾选: %s", ", ".join(PURCHASE_STATUS_WANTED))
        logger.info("     选择完成后回到终端按回车继续...")
        try:
            input()
        except (EOFError, OSError):
            pass
        logger.info("  状态选择已确认")
    else:
        logger.info("  跳过状态选择（PURCHASE_STATUS_WANTED 为空）")

    # ── 5. 勾选导出 EXCEL 设置 ────────────────────────────────────────
    _check_export_settings(wb)

    logger.info("✔ 查询条件填写完成")


def _fill_date_row(wb: WebBridgeClient, row_label: str,
                   start_date: str, end_date: str):
    """填写指定标签行（如 "创建日期"、"要求到货日期"）的日期范围。"""
    _set_one_date(wb, row_label, "开始", start_date)
    time.sleep(0.3)
    _set_one_date(wb, row_label, "结束", end_date)
    time.sleep(DELAY_SHORT)


def _set_one_date(wb: WebBridgeClient, row_label: str,
                  placeholder_keyword: str, date_str: str):
    """在包含 row_label 文字的 .next-row 中，设置匹配 placeholder 的日期 input。"""
    wb.evaluate(f"""
        (() => {{
            const rows = document.querySelectorAll('.next-row');
            let row = null;
            for (const r of rows) {{
                if (r.textContent.includes({_js_str(row_label)})) {{
                    row = r; break;
                }}
            }}
            if (!row) return 'row not found: ' + {_js_str(row_label)};
            const inputs = row.querySelectorAll('input[type="text"]');
            let el = null;
            for (const inp of inputs) {{
                if (inp.placeholder.includes({_js_str(placeholder_keyword)})) {{
                    el = inp; break;
                }}
            }}
            if (!el) return 'input not found';
            const ns = Object.getOwnPropertyDescriptor(
                HTMLInputElement.prototype, 'value'
            ).set;
            el.focus(); el.click();
            ns.call(el, {_js_str(date_str)});
            el.dispatchEvent(new Event('input', {{ bubbles: true }}));
            el.dispatchEvent(new Event('change', {{ bubbles: true }}));
            el.dispatchEvent(new FocusEvent('blur'));
            return 'ok';
        }})()
    """)


def _select_purchase_status(wb: WebBridgeClient):
    """打开采购单状态多选下拉，勾选 PURCHASE_STATUS_WANTED 中指定的状态。

    策略：先取消全部已选 → 再逐个勾选目标状态。
    状态组件为 hippo-select-multiple，下拉选项通过 React 渲染在 overlay 中。
    """
    wanted_json = json.dumps(list(PURCHASE_STATUS_WANTED), ensure_ascii=False)

    # 1. 点击状态触发区域，打开下拉
    wb.evaluate("""
        (() => {
            const trigger = document.querySelector('.hippo-select-multiple');
            if (trigger) {
                trigger.dispatchEvent(new MouseEvent('mousedown', {bubbles: true}));
                trigger.dispatchEvent(new MouseEvent('mouseup', {bubbles: true}));
                trigger.dispatchEvent(new MouseEvent('click', {bubbles: true}));
                return 'trigger clicked';
            }
            return 'trigger not found';
        })()
    """)
    time.sleep(DELAY_MEDIUM)

    # 2. 尝试多种方式查找并操作下拉选项
    result = wb.evaluate(f"""
        (() => {{
            const wanted = {wanted_json};
            const clicked = [];

            // 方式 A: 在 overlay 中查找 menu item
            const overlays = document.querySelectorAll('.next-overlay-wrapper');
            for (const ov of overlays) {{
                const items = ov.querySelectorAll('.next-menu-item, li, .next-checkbox-wrapper');
                for (const item of items) {{
                    const text = item.textContent.trim();
                    for (const w of wanted) {{
                        if (text.includes(w)) {{
                            const cb = item.querySelector('input[type="checkbox"]');
                            if (cb) {{
                                if (!cb.checked) {{ cb.click(); clicked.push(w); }}
                                else {{ clicked.push(w + '(already checked)'); }}
                            }} else {{
                                item.click();
                                clicked.push(w + '(click)');
                            }}
                            break;
                        }}
                    }}
                }}
                if (clicked.length > 0) break;
            }}

            // 方式 B: 在整个 document 中查找状态相关 checkbox
            if (clicked.length === 0) {{
                const allCbs = document.querySelectorAll('input[type="checkbox"]');
                for (const cb of allCbs) {{
                    const parentText = (cb.closest('label')?.textContent || '').trim();
                    for (const w of wanted) {{
                        if (parentText.includes(w)) {{
                            if (!cb.checked) {{ cb.click(); clicked.push(w); }}
                            break;
                        }}
                    }}
                }}
            }}

            return clicked.length > 0
                ? 'clicked: ' + JSON.stringify(clicked)
                : 'no items matched in any approach';
        }})()
    """)
    logger.info("  状态选择结果: %s", result)

    # 3. 关闭下拉
    wb.evaluate("""
        document.activeElement?.dispatchEvent(
            new KeyboardEvent('keydown', {key: 'Escape', bubbles: true})
        );
    """)
    time.sleep(DELAY_SHORT)


def _dismiss_dialogs(wb: WebBridgeClient):
    """关闭页面上的弹窗/对话框。只点 × 关闭链接，不点任何文字按钮。"""
    result = wb.evaluate("""
        (() => {
            // 只找关闭链接/图标，避免误点「去确认」等按钮
            const closeSel = 'a.next-dialog-close, .next-dialog-close, ' +
                '[class*="dialog"] [class*="close"], .next-overlay-wrapper .next-dialog-close';
            const closeBtns = document.querySelectorAll(closeSel);
            for (const btn of closeBtns) {
                if (btn.offsetParent !== null) {
                    btn.click();
                    return 'closed: ' + (btn.className?.substring(0, 30) || btn.tagName);
                }
            }
            return 'no dialog found';
        })()
    """)
    if 'closed' in str(result):
        logger.info("  已关闭弹窗: %s", result)
        time.sleep(DELAY_SHORT)


def _check_export_settings(wb: WebBridgeClient):
    """勾选导出 EXCEL 设置中的「越库类型订单导出时带有门店配货信息或仓的调拨信息」。"""
    logger.info("  勾选导出EXCEL设置...")
    result = wb.evaluate("""
        (() => {
            // 找到包含"导出EXCEL设置"标签的那一行
            const rows = document.querySelectorAll('.next-row');
            let targetRow = null;
            for (const r of rows) {
                if (r.textContent.includes('导出EXCEL设置')) { targetRow = r; break; }
            }
            if (!targetRow) return 'row not found';

            const cbs = targetRow.querySelectorAll('input[type="checkbox"]');
            for (const cb of cbs) {
                const parentLabel = cb.closest('label');
                const parentText = parentLabel?.textContent?.trim() || '';
                const labelSpan = cb.parentElement?.querySelector('.next-checkbox-label');
                const labelText = labelSpan?.textContent?.trim() || '';

                if (labelText.includes('门店配货信息') || labelText.includes('仓的调拨信息') ||
                    parentText.includes('门店配货信息')) {
                    if (parentLabel) {
                        parentLabel.click();
                        return 'clicked via label, checked=' + cb.checked;
                    }
                    cb.click();
                    return 'clicked via input, checked=' + cb.checked;
                }
            }
            const found = Array.from(cbs).map(cb => {
                const ls = cb.parentElement?.querySelector('.next-checkbox-label');
                return ls?.textContent?.trim()?.substring(0, 40) || '?';
            });
            return 'not found. labels: ' + JSON.stringify(found);
        })()
    """)
    logger.info("  导出设置: %s", result)
    time.sleep(DELAY_SHORT)


def click_search(wb: WebBridgeClient):
    """点击「查询」按钮（按文字匹配），等待表格数据加载完成。"""
    logger.info("▶ 点击查询...")
    wb.evaluate("""
        (() => {
            const btns = document.querySelectorAll('button');
            for (const b of btns) {
                if (b.textContent.trim() === '查询') {
                    b.click();
                    return 'clicked';
                }
            }
            return 'not found';
        })()
    """)
    # 等待表格加载
    for _ in range(20):
        time.sleep(0.5)
        rows = wb.evaluate(
            "document.querySelectorAll('.next-table-body tr').length"
        )
        if rows > 0:
            break
    time.sleep(1)
    logger.info("✔ 查询已触发，结果已加载")


def get_total_pages(wb: WebBridgeClient) -> int:
    """逐页点击「下一页」直到按钮 disabled，找出真实总页数。"""
    total_pages = 1
    for attempt in range(500):
        # 检查「下一页」按钮是否 disabled
        # 注意：所有按钮 class 都有 "next-btn"（Next UI 框架前缀），
        # 必须用 classList 精确匹配 "next"，不能用 includes。
        can_next = wb.evaluate("""
            (() => {
                const pager = document.querySelector('.next-pagination');
                if (!pager) return false;
                const btns = pager.querySelectorAll('button');
                for (const b of btns) {
                    if (b.classList.contains('next')) {
                        if (b.disabled || b.hasAttribute('disabled') ||
                            b.classList.contains('disabled')) return false;
                        return true;
                    }
                }
                return false;
            })()
        """)
        if not can_next:
            break

        # 点击 next 箭头（用 classList，不用 includes）
        wb.evaluate("""
            (() => {
                const pager = document.querySelector('.next-pagination');
                const btns = pager.querySelectorAll('button');
                for (const b of btns) {
                    if (b.classList.contains('next')) { b.click(); return 'ok'; }
                }
                return 'nf';
            })()
        """)
        time.sleep(0.8)
        total_pages += 1

    # 读取当前页（带 current class 的按钮）
    current = wb.evaluate("""
        (() => {
            const pager = document.querySelector('.next-pagination');
            if (!pager) return '1';
            const btns = pager.querySelectorAll('button');
            for (const b of btns) {
                if (b.className.includes('current')) {
                    return b.textContent.trim();
                }
            }
            return String('?');
        })()
    """)
    final_pages = int(str(current).strip()) if str(current).strip().isdigit() else total_pages

    logger.info("  ===== 查询结果：共 %d 页 =====", final_pages)

    # 回到第一页
    wb.evaluate("""
        (() => {
            const pager = document.querySelector('.next-pagination');
            if (!pager) return;
            const btns = pager.querySelectorAll('button');
            for (const b of btns) {
                if (b.textContent.trim() === '1') { b.click(); return; }
            }
        })()
    """)
    time.sleep(1.5)
    return max(final_pages, 1)


def export_current_page(wb: WebBridgeClient, page_num: int,
                        download_dir: str) -> str | None:
    """全选当前页 → 导出 Excel → 等待下载。文件留在 download_dir 中。

    导出后关闭可能弹出的新 tab（如退货确认页面）。
    """
    logger.info("  ▶ 第 %d 页：全选并导出...", page_num)

    # ── 全选 ─────────────────────────────────────────────────────────
    try:
        wb.evaluate("""
            (() => {
                const headerCbs = document.querySelectorAll(
                    '.next-table-header input[type="checkbox"]'
                );
                for (const cb of headerCbs) {
                    const label = cb.closest('label');
                    if (label) {
                        if (cb.checked) { label.click(); }
                        label.click();
                        return 'clicked label, cb checked=' + cb.checked;
                    }
                }
                return 'no label found';
            })()
        """)
        time.sleep(DELAY_SHORT)
    except Exception as e:
        logger.warning("  全选失败: %s", e)

    # 记录下载前的文件和 tab 数
    before = set(os.listdir(download_dir))
    tabs_before = len(wb.list_tabs())

    # ── 点击导出 Excel ───────────────────────────────────────────────
    wb.evaluate("""
        (() => {
            const btns = document.querySelectorAll('button');
            for (const b of btns) {
                if (b.textContent.includes('导出Excel')) {
                    b.click();
                    return 'clicked';
                }
            }
            return 'not found';
        })()
    """)
    logger.info("  等待第 %d 页文件下载...", page_num)

    # 等待新文件出现
    new_file = wait_for_new_file(download_dir, before, timeout=90)
    if new_file:
        logger.info("  ✔ 已下载: %s", new_file)
    else:
        logger.error("  ✗ 第 %d 页下载失败或超时", page_num)

    # ── 关闭导出时弹出的新 tab ───────────────────────────────────────
    time.sleep(1)
    tabs_after = wb.list_tabs()
    new_tabs = len(tabs_after) - tabs_before
    if new_tabs > 0:
        logger.info("  检测到 %d 个新 tab，正在关闭...", new_tabs)
        # 切回当前 session 的 tab，然后关闭多余的
        wb.find_tab("https://portalpro.hemaos.com/pages/supplierPlatformNew/purchaseList.html")
        # 关闭 popup tab（通过 evaluate 在当前 context 中关不掉其他 tab，
        # 用 close_tab 关当前 tab，但我们需要先切到多余 tab）
        for t in tabs_after:
            url = t.get('url', '')
            if 'return' in url.lower() or '退货' in url or 'confirm' in url.lower():
                try:
                    wb.find_tab(url)
                    wb.evaluate("window.close()")
                    logger.info("  已尝试关闭: %s", url[:60])
                except Exception:
                    pass
        # 回到主 tab
        wb.find_tab("https://portalpro.hemaos.com/pages/supplierPlatformNew/purchaseList.html")

    return new_file


def go_to_next_page(wb: WebBridgeClient, current: int) -> bool:
    """翻到下一页。重试 3 次（每次间隔递增），应对页面渲染延迟。"""
    target_page = current + 1
    for attempt in range(3):
        wait = 1 + attempt * 2
        time.sleep(wait)
        clicked = wb.evaluate(f"""
            (() => {{
                const btns = document.querySelectorAll('.next-pagination button');
                let found = [];
                for (const b of btns) {{
                    const t = b.textContent.trim();
                    found.push(t);
                    if (t === '{target_page}' && !b.disabled) {{
                        b.click();
                        return 'page_btn';
                    }}
                }}
                // 输出找到的所有按钮文字，方便调试
                return 'btns: ' + JSON.stringify(found);
            }})()
        """)
        if clicked == 'page_btn':
            time.sleep(DELAY_LONG)
            return True
        # 备用：箭头按钮
        clicked2 = wb.evaluate("""
            (() => {
                const nextBtn = document.querySelector(
                    '.next-pagination-next:not(.disabled), ' +
                    'button[aria-label="下一页"]:not([disabled])'
                );
                if (nextBtn) { nextBtn.click(); return true; }
                return false;
            })()
        """)
        if clicked2:
            time.sleep(DELAY_LONG)
            return True
        logger.warning("  翻页重试 %d/3，按钮: %s", attempt + 1, clicked)
    logger.warning("  3 次重试后仍找不到第 %d 页按钮", target_page)
    return False


def run(start_date: str = None, end_date: str = None, add_days: int = 0):
    """主流程。

    日期优先级：CLI 参数 > config/settings.py > 自动计算（今天-7天 ~ 今天）
    --start + --add：end = start + add_days 天
    """
    setup_logging(os.path.join(os.path.dirname(__file__), "logs"))

    # 计算要求到货日期（优先级：CLI > config > 自动）
    if start_date:
        delivery_start = start_date
        if end_date:
            delivery_end = end_date
        else:
            dt = datetime.strptime(start_date, "%Y-%m-%d")
            dt_end = dt + timedelta(days=add_days if add_days else 6)
            delivery_end = dt_end.strftime("%Y-%m-%d")
    elif DELIVERY_DATE_START and DELIVERY_DATE_END:
        delivery_start = DELIVERY_DATE_START
        delivery_end = DELIVERY_DATE_END
    else:
        _, delivery_end = get_date_range(end_date, days_back=0)
        delivery_start, _ = get_date_range(delivery_end, days_back=7)

    create_start, _ = get_date_range(delivery_start, days_back=CREATE_OFFSET_DAYS)
    create_end = delivery_end

    logger.info("创建日期: %s ~ %s", create_start, create_end)
    logger.info("要求到货日期: %s ~ %s", delivery_start, delivery_end)

    # ── 创建 WebBridge 客户端、自动拉起 daemon ────────────────────────
    daemon_bin = os.path.expandvars(
        r"%USERPROFILE%\.kimi-webbridge\bin\kimi-webbridge.exe"
    )
    if not os.path.isfile(daemon_bin):
        logger.error("找不到 WebBridge daemon: %s", daemon_bin)
        logger.error("请确认已安装 Kimi WebBridge")
        return

    # 尝试启动 daemon（多次启动不影响，daemon 自带幂等）
    logger.info("正在启动 WebBridge daemon...")
    subprocess.run(
        [daemon_bin, "start"],
        capture_output=True,
        timeout=10,
        creationflags=subprocess.CREATE_NO_WINDOW,
    )
    time.sleep(2)

    try:
        wb = WebBridgeClient(SESSION, port=WEBBRIDGE_PORT)
    except WebBridgeError as e:
        logger.error("WebBridge daemon 启动失败: %s", e)
        return

    # 使用 Chrome 默认下载目录（WebBridge 通过真实 Chrome 下载，文件自动到 Downloads）
    actual_dl_dir = os.path.expandvars(r"%USERPROFILE%\Downloads")
    os.makedirs(actual_dl_dir, exist_ok=True)
    logger.info("下载目录: %s", actual_dl_dir)

    try:
        # ── 1. 导航到采购单列表（iframe 内页，直接访问） ────────────
        logger.info("▶ 打开采购单列表: %s", PURCHASE_LIST_URL)
        wb.navigate(PURCHASE_LIST_URL, new_tab=True, group_title="采购单下载")
        time.sleep(DELAY_LONG)

        # 关闭可能存在的弹窗（如「退货单确认提示」）
        _dismiss_dialogs(wb)

        # ── 2. 填写查询条件 ────────────────────────────────────────
        fill_search_form(wb, create_start, create_end,
                         delivery_start, delivery_end)

        # ── 3. 点击查询 ────────────────────────────────────────────
        click_search(wb)

        # ── 4. 获取总页数 ──────────────────────────────────────────
        time.sleep(DELAY_MEDIUM)
        total_pages = get_total_pages(wb)
        logger.info("共 %d 页数据，开始逐页下载...", total_pages)

        # ── 5. 逐页导出 ────────────────────────────────────────────
        downloaded = []
        for page in range(1, min(total_pages, MAX_PAGES) + 1):
            logger.info("━━━ 第 %d/%d 页 ━━━", page, total_pages)

            result = export_current_page(wb, page, actual_dl_dir)
            if result:
                downloaded.append(result)

            if page < total_pages:
                if not go_to_next_page(wb, page):
                    logger.info("已到最后一页，停止翻页")
                    break

        logger.info("\n完成！共下载 %d 个文件，保存至: %s",
                      len(downloaded), actual_dl_dir)
        for f in downloaded:
            logger.info(f"  {f}")

    except WebBridgeError as e:
        logger.exception("WebBridge 异常: %s", e)
    except Exception as e:
        logger.exception("脚本异常终止: %s", e)


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="盒马采购单自动下载")
    parser.add_argument("--start", help="要求到货日期 开始 (YYYY-MM-DD)")
    parser.add_argument("--end",   help="要求到货日期 结束 (YYYY-MM-DD)")
    parser.add_argument("--add",   type=int, default=0,
                        help="从 --start 往后加 N 天得到结束日期（与 --end 二选一）")
    args = parser.parse_args()
    run(start_date=args.start, end_date=args.end, add_days=args.add)
