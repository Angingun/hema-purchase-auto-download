"""
采购单自动下载脚本 (WebBridge 版)
网站: portalpro.hemaos.com
依赖: Kimi WebBridge daemon（端口由 settings.py 配置）+ Chrome 扩展
功能: 通过 WebBridge 操作真实浏览器，自动填写查询条件并逐页导出 Excel

前置条件:
  1. Chrome 已安装 Kimi WebBridge 扩展
  2. daemon 已启动（打开扩展面板即自动启动）
  3. Chrome 已登录盒马供应商平台
"""

import os
import shutil
import socket
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
    setup_logging, get_date_range, wait_download_complete,
    verify_excel_file, verify_export_total,
)
from utils.page_state import (
    get_page_state, wait_table_ready, wait_query_result_total, PageState,
)
from utils.webbridge_client import WebBridgeClient, WebBridgeError

logger = logging.getLogger(__name__)

SESSION = "hema-" + datetime.now().strftime("%H%M%S")


def _start_webbridge_daemon(timeout: int = 15) -> dict[str, object]:
    """执行一次幂等 start，并返回可用于诊断的结构化结果。"""
    daemon_bin = os.path.expandvars(
        r"%USERPROFILE%\.kimi-webbridge\bin\kimi-webbridge.exe"
    )
    result: dict[str, object] = {
        "ok": False,
        "binary": daemon_bin,
        "returncode": None,
        "stdout": "",
        "stderr": "",
        "error": None,
    }
    if not os.path.isfile(daemon_bin):
        result["error"] = f"找不到 WebBridge daemon: {daemon_bin}"
        return result

    if _wait_local_port(WEBBRIDGE_PORT, timeout=0.3):
        result.update({
            "ok": True,
            "stdout": f"daemon 已在 127.0.0.1:{WEBBRIDGE_PORT} 监听",
        })
        return result

    try:
        proc = subprocess.run(
            [
                daemon_bin,
                "start",
                "--addr",
                f"127.0.0.1:{WEBBRIDGE_PORT}",
            ],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
    except subprocess.TimeoutExpired:
        result["error"] = f"WebBridge start 超时（{timeout}s）"
        return result
    except OSError as exc:
        result["error"] = f"无法执行 WebBridge daemon: {exc}"
        return result

    result.update({
        "ok": proc.returncode == 0,
        "returncode": proc.returncode,
        "stdout": (proc.stdout or "").strip(),
        "stderr": (proc.stderr or "").strip(),
    })
    if proc.returncode != 0:
        result["error"] = f"WebBridge start 返回非 0: {proc.returncode}"
    return result


def _wait_local_port(port: int, timeout: float = 5.0) -> bool:
    """等待本地 TCP 端口开始监听。"""
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=0.5):
                return True
        except OSError:
            time.sleep(0.25)
    return False


def check_webbridge(timeout: float = 10.0) -> bool:
    """检查 daemon 监听与 Chrome 扩展握手，不执行页面操作。"""
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except (AttributeError, OSError):
        pass

    start_result = _start_webbridge_daemon()
    print(f"WebBridge 可执行文件: {start_result['binary']}")
    if start_result.get("stdout"):
        print(f"start: {start_result['stdout']}")
    if not start_result["ok"]:
        print(f"[警告] {start_result['error']}")
        if start_result.get("stderr"):
            print(f"stderr: {start_result['stderr']}")
        if not os.path.isfile(str(start_result["binary"])):
            return False

    port_wait = min(5.0, timeout)
    if not _wait_local_port(WEBBRIDGE_PORT, timeout=port_wait):
        start_state = (
            "启动命令已返回成功，但"
            if start_result["ok"]
            else "启动命令异常，且"
        )
        print(
            f"[失败] daemon {start_state}端口 {WEBBRIDGE_PORT} "
            "未监听；进程可能启动后立即退出。"
        )
        return False

    print(f"[通过] daemon 正在监听 127.0.0.1:{WEBBRIDGE_PORT}")
    wb = WebBridgeClient("webbridge-health-check", port=WEBBRIDGE_PORT)
    try:
        wb.wait_ready(timeout=max(1.0, timeout - port_wait))
        tab_count = len(wb.list_tabs())
    except WebBridgeError as exc:
        print(f"[失败] daemon 可访问，但 Chrome 扩展未完成握手: {exc}")
        return False

    print(f"[通过] Chrome 扩展已连接，可读取 {tab_count} 个标签页")
    return True


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

    # ── 4. 采购单状态（多选）────────────────────────────────────────
    if PURCHASE_STATUS_WANTED:
        status_result = select_purchase_statuses(
            wb, list(PURCHASE_STATUS_WANTED), fallback_manual=True
        )
        if not status_result["ok"]:
            raise RuntimeError(
                "采购单状态未通过校验，缺少: "
                + ", ".join(status_result["missing"])
                + "；多余: "
                + ", ".join(status_result["extra"])
            )
        logger.info(
            "  状态选择已校验（%s）: %s",
            status_result["method"],
            ", ".join(status_result["selected"]),
        )
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


def _read_selected_purchase_statuses(wb: WebBridgeClient) -> list[str]:
    """反读第一个多选组件中已选的采购单状态。"""
    selected = wb.evaluate("""
        (() => {
            const select = document.querySelectorAll('.hippo-select-multiple')[0];
            if (!select) return [];
            return Array.from(select.querySelectorAll(
                '.next-tag-deletable .next-tag-body'
            )).map(el => el.textContent.trim()).filter(Boolean);
        })()
    """)
    return list(dict.fromkeys(selected or []))


def select_purchase_statuses(
    wb: WebBridgeClient,
    wanted: list[str],
    timeout: int = 10,
    fallback_manual: bool = True,
) -> dict[str, object]:
    """将采购单状态设置为 wanted，并以页面已选标签作为最终校验。"""
    wanted = list(dict.fromkeys(status.strip() for status in wanted if status.strip()))
    options_seen: list[str] = []
    method = "automatic"
    auto_error: str | None = None

    try:
        selected = _read_selected_purchase_statuses(wb)
        if not wb.element_exists('.hippo-select-multiple input'):
            raise RuntimeError("未找到采购单状态多选输入框")

        for status in [value for value in selected if value not in wanted]:
            marked = wb.evaluate(f"""
                (() => {{
                    const select = document.querySelectorAll(
                        '.hippo-select-multiple'
                    )[0];
                    if (!select) return false;
                    for (const old of document.querySelectorAll(
                        '[data-status-remove-target]'
                    )) {{
                        old.removeAttribute('data-status-remove-target');
                    }}
                    for (const tag of select.querySelectorAll('.next-tag-deletable')) {{
                        const body = tag.querySelector('.next-tag-body');
                        if (body?.textContent.trim() === {_js_str(status)}) {{
                            tag.setAttribute('data-status-remove-target', 'true');
                            return true;
                        }}
                    }}
                    return false;
                }})()
            """)
            if not marked:
                raise RuntimeError(f"无法定位待移除状态: {status}")
            wb.click('[data-status-remove-target="true"] .next-tag-tail')
            time.sleep(DELAY_SHORT)

        selected = _read_selected_purchase_statuses(wb)
        for status in [value for value in wanted if value not in selected]:
            wb.fill('.hippo-select-multiple input', status)
            deadline = time.time() + timeout
            option_ready = False
            while time.time() < deadline:
                option_info = wb.evaluate(f"""
                    (() => {{
                        const seen = [];
                        let matched = false;
                        for (const old of document.querySelectorAll(
                            '[data-status-option-target]'
                        )) {{
                            old.removeAttribute('data-status-option-target');
                        }}
                        for (const item of document.querySelectorAll(
                            '.hippo-select-menu-item'
                        )) {{
                            if (item.offsetParent === null) continue;
                            const text = item.textContent.trim();
                            if (text) seen.push(text);
                            if (text === {_js_str(status)}) {{
                                item.setAttribute('data-status-option-target', 'true');
                                matched = true;
                            }}
                        }}
                        return {{matched, seen}};
                    }})()
                """) or {}
                options_seen.extend(option_info.get("seen", []))
                if option_info.get("matched"):
                    option_ready = True
                    break
                time.sleep(0.2)
            if not option_ready:
                raise TimeoutError(f"未出现采购单状态选项: {status}")
            wb.click('[data-status-option-target="true"]')
            time.sleep(DELAY_SHORT)

        selected = _read_selected_purchase_statuses(wb)
        missing = [value for value in wanted if value not in selected]
        extra = [value for value in selected if value not in wanted]
        if missing or extra:
            raise RuntimeError(
                f"自动选择后校验失败，missing={missing}, extra={extra}"
            )
    except Exception as exc:
        auto_error = str(exc)
        selected = _read_selected_purchase_statuses(wb)
        logger.warning("  自动选择采购单状态失败: %s", auto_error)

    missing = [value for value in wanted if value not in selected]
    extra = [value for value in selected if value not in wanted]
    if (missing or extra) and fallback_manual:
        method = "manual_fallback"
        logger.info("  请在浏览器中手动调整采购单状态：")
        logger.info("     需要选择: %s", ", ".join(wanted))
        logger.info("     调整完成后回到终端按回车继续...")
        try:
            input()
        except (EOFError, OSError):
            pass
        selected = _read_selected_purchase_statuses(wb)
        missing = [value for value in wanted if value not in selected]
        extra = [value for value in selected if value not in wanted]

    ok = not missing and not extra
    return {
        "ok": ok,
        "wanted": wanted,
        "selected": selected,
        "missing": missing,
        "extra": extra,
        "options_seen": list(dict.fromkeys(options_seen)),
        "method": method,
        "error": None if ok else auto_error or "采购单状态校验失败",
    }


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


def click_search(wb: WebBridgeClient) -> int:
    """点击「查询」按钮，等待本次查询的第一页数据替换旧结果。"""
    previous_state = get_page_state(wb)
    previous_order_hash = previous_state.get('order_hash')

    logger.info("▶ 点击查询...")
    clicked = wb.evaluate("""
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

    if clicked != 'clicked':
        raise RuntimeError(f"查询按钮点击失败: {clicked}")

    # 旧查询结果可能仍在 DOM 中，必须先确认本次查询的第一页数据已替换。
    wait_table_ready(
        wb,
        expected_page=1,
        previous_order_hash=str(previous_order_hash) if previous_order_hash else None,
    )
    query_total = wait_query_result_total(wb)
    logger.info("✔ 查询已触发，结果已加载，共 %d 条数据", query_total)
    return query_total

def get_total_pages(wb: WebBridgeClient) -> int:
    """逐页点击「下一页」直到按钮 disabled，找出真实总页数。"""
    logger.info("  Reading initial page state...")
    try:
        init_state = get_page_state(wb)
        logger.info("  Initial: page=%d rows=%d key=%s",
                    init_state['page_num'], init_state['row_count'],
                    init_state['first_row_key'])
    except Exception:
        pass
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
    try:
        back_state = get_page_state(wb)
        logger.info("  Back to page 1: rows=%d key=%s",
                    back_state['row_count'], back_state['first_row_key'])
    except Exception:
        pass
    return max(final_pages, 1)


def go_to_page(wb: WebBridgeClient, target_page: int,
               timeout: int = 30) -> PageState:
    """翻到目标页码并等待表格稳定。

    优先点击可见页码按钮；若不在视窗内，使用下一页按钮逐页前进。
    每次点击后调用 wait_table_ready()。
    无法到达目标页时抛出异常。
    """
    current = get_page_state(wb)
    current_page_num = current.get('page_num', 0)

    logger.info("  Navigate from page %d to page %d",
                current_page_num, target_page)

    if current_page_num == target_page:
        return wait_table_ready(wb, target_page, timeout=timeout)

    if current_page_num > target_page:
        raise ValueError(
            f"Cannot go back: current page {current_page_num} "
            f"> target {target_page}"
        )

    # 逐页前进
    while current_page_num < target_page:
        expected_next = current_page_num + 1
        result = wb.evaluate(f"""
            (() => {{
                const pager = document.querySelector('.next-pagination');
                if (!pager) return 'no_pager';
                const btns = pager.querySelectorAll('button');
                for (const b of btns) {{
                    if (b.textContent.trim() === '{expected_next}'
                        && !b.disabled) {{
                        b.click();
                        return 'ok';
                    }}
                }}
                for (const b of btns) {{
                    if (b.classList.contains('next')
                        && !b.disabled
                        && !b.hasAttribute('disabled')) {{
                        b.click();
                        return 'next';
                    }}
                }}
                return 'stuck';
            }})()
        """)

        if result == 'stuck':
            raise RuntimeError(
                f"Pagination stuck at page {current_page_num} "
                f"(target {target_page})"
            )

        current_page_num += 1

        if current_page_num < target_page:
            time.sleep(1.0)
        else:
            return wait_table_ready(wb, target_page, timeout=timeout)

    return wait_table_ready(wb, target_page, timeout=timeout)



def go_to_next_page(wb: WebBridgeClient, current_state: PageState,
                    timeout: int = 30) -> PageState | None:
    """点击下一页，并等待页码与订单 hash 都变化。

    返回 None 表示当前页已经是最后一页。
    """
    if not current_state.get('can_next'):
        return None

    current_page = int(current_state.get('page_num', 0) or 0)
    current_hash = current_state.get('order_hash')
    clicked = wb.evaluate("""
        (() => {
            const pager = document.querySelector('.next-pagination');
            if (!pager) return 'no_pager';
            for (const b of pager.querySelectorAll('button')) {
                if (b.classList.contains('next') &&
                    !b.disabled &&
                    !b.hasAttribute('disabled') &&
                    !b.classList.contains('disabled')) {
                    b.click();
                    return 'clicked';
                }
            }
            return 'disabled';
        })()
    """)
    if clicked != 'clicked':
        logger.info("  Next page unavailable: %s", clicked)
        return None

    return wait_table_ready(
        wb,
        current_page + 1,
        timeout=timeout,
        previous_order_hash=str(current_hash) if current_hash else None,
    )

def export_current_page(wb: WebBridgeClient, page_num: int,
                        download_dir: str,
                        page_state: PageState | None = None) -> dict:
    """全选当前页 → 导出 Excel → 强校验下载。

    参数：
      wb: WebBridge client
      page_num: 期望页号
      download_dir: 监控的下载目录
      page_state: 预读的页面状态（来自 wait_table_ready）

    返回结构化结果（字段兼容汇总日志）：
      page_num, ok, ui_row_count, file_path, file_size,
      excel_row_count, error
    """
    result: dict = {
        'page_num': page_num,
        'ok': False,
        'ui_row_count': 0,
        'file_path': None,
        'file_size': None,
        'excel_row_count': None,
        'page_order_count': 0,
        'page_order_hash': None,
        'excel_order_count': 0,
        'excel_order_hash': None,
        'skipped': False,
        'error': None,
    }

    # ── 导出前校验（单元 7）─────────────────────────────────────────
    if page_state is None:
        page_state = get_page_state(wb)

    actual_page = page_state.get('page_num', 0)
    row_count = page_state.get('row_count', 0)
    first_key = page_state.get('first_row_key')

    result['ui_row_count'] = row_count
    result['page_order_count'] = len(page_state.get('order_ids', []))
    result['page_order_hash'] = page_state.get('order_hash')
    logger.info(
        "  Pre-export check: target_page=%d actual_page=%d rows=%d key=%s",
        page_num, actual_page, row_count, first_key,
    )

    if actual_page != page_num:
        result['error'] = (
            f"Page mismatch: expected {page_num}, actual {actual_page}"
        )
        logger.error("  ✗ %s", result['error'])
        return result

    if row_count == 0:
        logger.warning("  Page %d has 0 visible rows, skipping export",
                      page_num)
        result['skipped'] = True
        result['error'] = 'empty_page'
        return result

    # ── 全选 ─────────────────────────────────────────────────────────
    logger.info("  ▶ 全选当前页...")
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
                        return 'ok';
                    }
                }
                return 'no_label';
            })()
        """)
        time.sleep(DELAY_SHORT)
    except Exception as e:
        logger.warning("  全选异常: %s", e)

    # ── 记录下载前状态 ──────────────────────────────────────────────
    before_files = set(os.listdir(download_dir))
    tabs_before = len(wb.list_tabs())

    # ── 点击导出 Excel ──────────────────────────────────────────────
    wb.evaluate("""
        (() => {
            const btns = document.querySelectorAll('button');
            for (const b of btns) {
                if (b.textContent.includes('导出Excel')) {
                    b.click();
                    return 'clicked';
                }
            }
            return 'not_found';
        })()
    """)
    logger.info("  等待第 %d 页文件下载...", page_num)

    # ── 强校验下载（单元 8）──────────────────────────────────────────
    dl_result = wait_download_complete(
        download_dir, before_files, timeout=DELAY_DOWNLOAD + 90
    )

    if dl_result['ok']:
        result['file_path'] = dl_result['path']
        result['file_size'] = dl_result['size_bytes']
        logger.info("  Download complete: %s (%d bytes)",
                    dl_result['filename'], dl_result['size_bytes'])

        excel_result = verify_excel_file(str(dl_result['path']))
        result['excel_row_count'] = excel_result.get('row_count')
        result['excel_order_count'] = len(excel_result.get('order_ids', []))
        result['excel_order_hash'] = excel_result.get('order_hash')
        if excel_result['ok'] and excel_result.get('order_hash') == page_state.get('order_hash'):
            result['ok'] = True
            logger.info(
                "  Excel validation passed: sheets=%s rows=%s cols=%s orders=%s",
                excel_result.get('sheet_names'),
                excel_result.get('row_count'),
                excel_result.get('column_count'),
                len(excel_result.get('order_ids', [])),
            )
        else:
            if excel_result['ok']:
                result['error'] = (
                    'Excel order hash mismatch: '
                    f"page={str(page_state.get('order_hash'))[:12]} "
                    f"excel={str(excel_result.get('order_hash'))[:12]}"
                )
            else:
                result['error'] = (
                    'Excel validation failed: '
                    f"{excel_result.get('error')}"
                )
            logger.error("  %s", result['error'])
    else:
        result['error'] = dl_result['error']
        logger.error("  ✗ 下载失败: %s", dl_result['error'])

    # ── 关闭导出弹出新 tab ─────────────────────────────────────────
    time.sleep(1)
    tabs_after = wb.list_tabs()
    new_tabs = len(tabs_after) - tabs_before
    if new_tabs > 0:
        logger.info("  检测到 %d 个新 tab，正在关闭...", new_tabs)
        for t in tabs_after:
            url = t.get('url', '')
            if ('return' in url.lower() or '退货' in url
                    or 'confirm' in url.lower()):
                try:
                    wb.find_tab(url)
                    wb.evaluate("window.close()")
                except Exception:
                    pass
        wb.find_tab(
            "https://portalpro.hemaos.com/pages/supplierPlatformNew/purchaseList.html"
        )

    return result


def _move_verified_export(result: dict, destination_dir: str) -> dict:
    """将校验成功的下载文件移动到归档目录，不覆盖同名文件。"""
    if not result.get('ok') or not result.get('file_path'):
        return result

    source = os.path.abspath(str(result['file_path']))
    destination_dir = os.path.abspath(os.path.expandvars(destination_dir))

    try:
        os.makedirs(destination_dir, exist_ok=True)
        destination = os.path.join(destination_dir, os.path.basename(source))
        if os.path.normcase(source) == os.path.normcase(destination):
            return result

        stem, suffix = os.path.splitext(destination)
        candidate = destination
        sequence = 1
        while os.path.exists(candidate):
            candidate = f"{stem} ({sequence}){suffix}"
            sequence += 1

        moved_path = shutil.move(source, candidate)
    except OSError as exc:
        result['ok'] = False
        result['error'] = f"Failed to move verified export: {exc}"
        logger.error("  移动已校验文件失败: %s", exc)
        return result

    result['file_path'] = os.path.abspath(moved_path)
    logger.info("  已移动到归档目录: %s", result['file_path'])
    return result

def _print_summary(results: list[dict], download_dir: str,
                   total_validation: dict | None = None):
    """输出运行汇总日志。"""
    success = [r for r in results if r.get('ok')]
    skipped = [r for r in results if r.get('skipped')]
    failed = [r for r in results if not r.get('ok') and not r.get('skipped')]
    total_ui_rows = sum(r.get('ui_row_count', 0) or 0 for r in results)
    total_excel_rows = sum(
        r.get('excel_row_count', 0) or 0
        for r in results if r.get('excel_row_count')
    )

    logger.info("=" * 50)
    logger.info("  运行汇总")
    logger.info("=" * 50)
    logger.info("  总处理页数: %d", len(results))
    logger.info("  成功页数:   %d", len(success))
    logger.info("  Skipped pages: %d", len(skipped))
    logger.info("  失败页数:   %d", len(failed))
    logger.info("  UI 可见行数合计: %d", total_ui_rows)
    logger.info("  Excel 行数合计: %d", total_excel_rows)
    if total_validation is not None:
        logger.info("  查询结果汇总: %d", total_validation['expected_total'])
        logger.info("  Excel 唯一采购单数: %d", total_validation['actual_total'])
        logger.info("  总量差额: %+d", total_validation['difference'])
        logger.info("  总量核验: %s", "通过" if total_validation['ok'] else "失败")
        if total_validation.get('duplicate_count'):
            logger.warning("  跨文件重复采购单数: %d", total_validation['duplicate_count'])
        if total_validation.get('error'):
            logger.error("  总量核验原因: %s", total_validation['error'])
    if failed:
        logger.info("  失败详情:")
        for r in failed:
            logger.info("    - 第 %d 页: %s",
                       r['page_num'], r.get('error', '未知'))
    if skipped:
        logger.info("  Skipped details:")
        for r in skipped:
            logger.info("    - Page %d: %s",
                       r['page_num'], r.get('error', 'unknown'))
    if success:
        logger.info("  下载文件目录: %s", download_dir)
        for r in success:
            path = r.get('file_path', '')
            size = r.get('file_size', 0)
            excel_rows = r.get('excel_row_count', '?')
            logger.info("    %s (%d bytes, excel=%s rows)",
                       os.path.basename(path) if path else '?',
                       size or 0, excel_rows)
    logger.info("=" * 50)


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
    # 尝试启动 daemon（多次启动不影响，daemon 自带幂等）
    logger.info("正在启动 WebBridge daemon...")
    start_result = _start_webbridge_daemon()
    if not start_result["ok"]:
        logger.warning("%s", start_result["error"])
        if start_result.get("stderr"):
            logger.warning("WebBridge start stderr: %s", start_result["stderr"])
        if not os.path.isfile(str(start_result["binary"])):
            return False
    if start_result.get("stdout"):
        logger.info("WebBridge start: %s", str(start_result["stdout"])[:300])
    wb = WebBridgeClient(SESSION, port=WEBBRIDGE_PORT)
    try:
        logger.info("等待 WebBridge daemon 与 Chrome 扩展就绪...")
        wb.wait_ready(timeout=30)
        logger.info("WebBridge 已就绪")
    except WebBridgeError as e:
        logger.error("WebBridge 未就绪: %s", e)
        logger.error("如果是重启电脑后的首次运行，请先打开 Chrome 并点击 Kimi WebBridge 扩展面板确认连接状态。")
        return

    actual_dl_dir = os.path.abspath(os.path.expandvars(CHROME_DOWNLOADS_DIR))
    archive_dir = os.path.abspath(os.path.expandvars(DOWNLOAD_DIR))
    os.makedirs(actual_dl_dir, exist_ok=True)
    logger.info("Chrome 下载监控目录: %s", actual_dl_dir)
    logger.info("校验后归档目录: %s", archive_dir)

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
        query_total = click_search(wb)

        # ── 4. 边导出边翻页，直到下一页禁用 ───────────────────────
        time.sleep(DELAY_MEDIUM)
        results: list[dict] = []
        seen_page_hashes: dict[str, int] = {}
        page = 1
        try:
            state = wait_table_ready(wb, page)
        except TimeoutError as e:
            logger.error("Page %d did not stabilize: %s", page, e)
            results.append({
                'page_num': page,
                'ok': False,
                'ui_row_count': 0,
                'file_path': None,
                'file_size': None,
                'excel_row_count': None,
                'page_order_count': 0,
                'page_order_hash': None,
                'excel_order_count': 0,
                'excel_order_hash': None,
                'skipped': False,
                'error': f'Stabilization failed: {e}',
            })
            state = None

        while state is not None and page <= MAX_PAGES:
            logger.info("━━━ 第 %d 页 ━━━", page)
            page_hash = state.get('order_hash')
            if page_hash:
                if page_hash in seen_page_hashes:
                    prev = seen_page_hashes[str(page_hash)]
                    msg = f"Duplicate page data: page {page} matches page {prev}"
                    logger.error(msg)
                    results.append({
                        'page_num': page,
                        'ok': False,
                        'ui_row_count': state.get('row_count', 0),
                        'file_path': None,
                        'file_size': None,
                        'excel_row_count': None,
                        'page_order_count': len(state.get('order_ids', [])),
                        'page_order_hash': page_hash,
                        'excel_order_count': 0,
                        'excel_order_hash': None,
                        'skipped': False,
                        'error': msg,
                    })
                    break
                seen_page_hashes[str(page_hash)] = page

            r = export_current_page(wb, page, actual_dl_dir, page_state=state)
            r = _move_verified_export(r, archive_dir)
            results.append(r)
            if not r.get("ok"):
                logger.error("Stopping after failed export on page %d: %s",
                             page, r.get('error'))
                break

            if not state.get('can_next'):
                logger.info("已到最后一页，停止翻页")
                break

            try:
                next_state = go_to_next_page(wb, state)
            except TimeoutError as e:
                logger.error("Failed to reach next page after page %d: %s", page, e)
                results.append({
                    'page_num': page + 1,
                    'ok': False,
                    'ui_row_count': 0,
                    'file_path': None,
                    'file_size': None,
                    'excel_row_count': None,
                    'page_order_count': 0,
                    'page_order_hash': None,
                    'excel_order_count': 0,
                    'excel_order_hash': None,
                    'skipped': False,
                    'error': f'Navigation failed: {e}',
                })
                break
            if next_state is None:
                logger.info("已到最后一页，停止翻页")
                break
            page += 1
            state = next_state

        if page > MAX_PAGES:
            logger.warning("达到 MAX_PAGES=%d，停止", MAX_PAGES)

        # ── 5. 运行汇总 ────────────────────────────────────────────
        total_validation = verify_export_total(results, query_total)
        _print_summary(results, archive_dir, total_validation)
        if not total_validation['ok']:
            logger.error("本次下载未通过查询总量核验")
            return False
        return True

    except WebBridgeError as e:
        logger.exception("WebBridge 异常: %s", e)
        return False
    except Exception as e:
        logger.exception("脚本异常终止: %s", e)
        return False


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="盒马采购单自动下载")
    parser.add_argument("--start", help="要求到货日期 开始 (YYYY-MM-DD)")
    parser.add_argument("--end",   help="要求到货日期 结束 (YYYY-MM-DD)")
    parser.add_argument("--add",   type=int, default=0,
                        help="从 --start 往后加 N 天得到结束日期（与 --end 二选一）")
    parser.add_argument(
        "--check-webbridge",
        action="store_true",
        help="只检查 WebBridge daemon 与 Chrome 扩展连接，不执行下载",
    )
    args = parser.parse_args()
    if args.check_webbridge:
        raise SystemExit(0 if check_webbridge() else 1)
    run(start_date=args.start, end_date=args.end, add_days=args.add)

