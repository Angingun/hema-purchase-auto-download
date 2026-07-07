"""页面状态读取和表格稳定性等待工具。"""

import hashlib
import logging
import time

from utils.webbridge_client import WebBridgeClient

logger = logging.getLogger(__name__)

PageState = dict[str, object]
"""字段约定：
    page_num: int
    row_count: int              -- 主表 body 的真实业务行数
    loading: bool
    first_row_key: str | None
    order_ids: list[str]        -- 当前页主表采购单号集合（保留页面顺序、去重）
    order_hash: str | None      -- order_ids 的 sha256 hash
    can_next: bool
    raw: dict
"""


def _hash_order_ids(order_ids: list[str]) -> str | None:
    if not order_ids:
        return None
    return hashlib.sha256("\n".join(sorted(set(order_ids))).encode("utf-8")).hexdigest()


def get_page_state(wb: WebBridgeClient) -> PageState:
    """读取当前页面状态，不修改 DOM。"""
    result = wb.evaluate("""
        (() => {
            const isVisible = (el) => {
                const style = getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                return el.offsetParent !== null &&
                    style.display !== 'none' &&
                    style.visibility !== 'hidden' &&
                    style.opacity !== '0' &&
                    rect.width > 0 && rect.height > 0;
            };

            const pager = document.querySelector('.next-pagination');
            let pageNum = 0;
            let pagerText = '';
            let canNext = false;
            const pagerButtons = [];
            if (pager) {
                pagerText = pager.textContent.trim().substring(0, 200);
                for (const b of pager.querySelectorAll('button')) {
                    const text = b.textContent.trim();
                    const disabled = b.disabled || b.hasAttribute('disabled') ||
                        b.classList.contains('disabled');
                    const cls = String(b.className || '');
                    pagerButtons.push({text, cls, disabled});
                    if (b.classList.contains('current') || cls.includes('current')) {
                        if (text && !isNaN(text)) pageNum = parseInt(text, 10);
                    }
                    if (b.classList.contains('next')) canNext = !disabled;
                }
            }

            const bodies = Array.from(document.querySelectorAll('.next-table-body'))
                .filter(isVisible)
                .map((body, index) => {
                    const rect = body.getBoundingClientRect();
                    const rows = Array.from(body.querySelectorAll('tr'));
                    return {body, index, width: rect.width, rows};
                })
                .filter(item => item.rows.length > 0)
                .sort((a, b) => b.width - a.width);
            const main = bodies.length ? bodies[0] : null;
            const rows = main ? main.rows : [];

            const orderIds = [];
            const seen = new Set();
            let firstRowKey = null;
            for (const row of rows) {
                const texts = Array.from(row.querySelectorAll('td, th'))
                    .map(cell => cell.textContent.trim())
                    .filter(Boolean);
                if (firstRowKey === null && texts.length > 0) {
                    firstRowKey = texts.slice(0, 5).join('|').substring(0, 200);
                }
                const orderId = texts.find(t => /^HPO/.test(t));
                if (orderId && !seen.has(orderId)) {
                    seen.add(orderId);
                    orderIds.push(orderId);
                }
            }

            const loadingCandidates = Array.from(document.querySelectorAll(
                '.next-table-loading:not(.next-hidden), ' +
                '.next-loading-indicator:not(.next-hidden), ' +
                '.next-loading-tip:not(.next-hidden), ' +
                '.next-icon-loading:not(.next-hidden)'
            ));
            const loading = loadingCandidates.some(isVisible);

            return {
                pageNum,
                rowCount: rows.length,
                loading,
                firstRowKey,
                orderIds,
                canNext,
                pagerText,
                pagerButtons,
                bodyCount: document.querySelectorAll('.next-table-body').length,
                visibleBodyCount: bodies.length,
                mainBodyIndex: main ? main.index : null,
                mainBodyWidth: main ? main.width : 0,
            };
        })()
    """)

    if not isinstance(result, dict):
        logger.warning("get_page_state: unexpected response type=%s val=%s",
                       type(result).__name__, result)
        return {
            'page_num': 0,
            'row_count': 0,
            'loading': False,
            'first_row_key': None,
            'order_ids': [],
            'order_hash': None,
            'can_next': False,
            'raw': {'error': f'unexpected response: {result}'},
        }

    order_ids = [str(v) for v in result.get('orderIds', []) if v]
    return {
        'page_num': int(result.get('pageNum', 0)),
        'row_count': int(result.get('rowCount', 0)),
        'loading': bool(result.get('loading', False)),
        'first_row_key': result.get('firstRowKey'),
        'order_ids': order_ids,
        'order_hash': _hash_order_ids(order_ids),
        'can_next': bool(result.get('canNext', False)),
        'raw': {
            'pager_text': result.get('pagerText', ''),
            'pager_buttons': result.get('pagerButtons', []),
            'body_count': result.get('bodyCount'),
            'visible_body_count': result.get('visibleBodyCount'),
            'main_body_index': result.get('mainBodyIndex'),
            'main_body_width': result.get('mainBodyWidth'),
        },
    }


def wait_table_ready(wb: WebBridgeClient, expected_page: int,
                     timeout: int = 30,
                     previous_order_hash: str | None = None) -> PageState:
    """等待指定页码的主表数据稳定。

    若 previous_order_hash 不为空，还要求当前页采购单号 hash 与前一页不同。
    """
    deadline = time.time() + timeout
    prev_state: PageState | None = None
    stable_count = 0

    while time.time() < deadline:
        state = get_page_state(wb)
        page_ok = state['page_num'] == expected_page
        loading_done = not state['loading']
        has_rows = int(state['row_count'] or 0) > 0
        hash_changed = (
            previous_order_hash is None or
            state.get('order_hash') != previous_order_hash
        )

        logger.debug(
            "wait_table_ready(p=%d): cur=%d rows=%d loading=%s hash=%s key=%s",
            expected_page, state['page_num'], state['row_count'],
            state['loading'], state.get('order_hash'), state['first_row_key'],
        )

        if page_ok and loading_done and has_rows and hash_changed:
            if prev_state is not None:
                if (prev_state['row_count'] == state['row_count'] and
                        prev_state['first_row_key'] == state['first_row_key'] and
                        prev_state.get('order_hash') == state.get('order_hash')):
                    stable_count += 1
                    if stable_count >= 2:
                        logger.info(
                            "Table stable on page %d: rows=%d orders=%d hash=%s key=%s",
                            state['page_num'], state['row_count'],
                            len(state.get('order_ids', [])),
                            str(state.get('order_hash'))[:12],
                            state['first_row_key'],
                        )
                        return state
                else:
                    stable_count = 0
            prev_state = state
        else:
            stable_count = 0
            prev_state = state

        remaining = deadline - time.time()
        if remaining > 0:
            time.sleep(min(0.5, remaining))

    final = get_page_state(wb)
    msg = (
        f"Table did not stabilize on page {expected_page} within {timeout}s. "
        f"Final: page={final['page_num']}, rows={final['row_count']}, "
        f"loading={final['loading']}, hash={final.get('order_hash')}, "
        f"can_next={final.get('can_next')}, key={final['first_row_key']}"
    )
    logger.error(msg)
    raise TimeoutError(msg)

