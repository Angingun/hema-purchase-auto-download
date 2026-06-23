"""Kimi WebBridge daemon HTTP API 封装。

WebBridge daemon 运行在 localhost:10086，接收 JSON 命令，
通过 CDP 操作真实 Chrome 浏览器。
"""

import time
import json
import logging
from typing import Any

import requests

logger = logging.getLogger(__name__)


class WebBridgeError(Exception):
    """WebBridge 命令执行失败"""


class WebBridgeClient:
    """封装 WebBridge daemon 的 HTTP REST API。

    用法:
        wb = WebBridgeClient("hema-download")
        wb.navigate("https://example.com", group_title="采购单下载")
        wb.fill("@e1", "some text")
        wb.click("button.submit")
        tree = wb.snapshot()
    """

    def __init__(self, session: str, port: int = 10086):
        self.session = session
        self.base_url = f"http://127.0.0.1:{port}"

    # ── 内部方法 ──────────────────────────────────────────────────────

    def _post(self, action: str, args: dict | None = None,
              timeout: int = 30) -> dict:
        """发送命令到 daemon，返回 data 字段。出错抛 WebBridgeError。"""
        body: dict[str, Any] = {
            "action": action,
            "args": args or {},
            "session": self.session,
        }
        try:
            resp = requests.post(
                f"{self.base_url}/command",
                json=body,
                timeout=timeout,
            )
            resp.encoding = "utf-8"
            result = resp.json()
        except requests.ConnectionError:
            raise WebBridgeError(
                "无法连接到 WebBridge daemon。请确认 daemon 已启动："
                r"& \"$env:USERPROFILE\.kimi-webbridge\bin\kimi-webbridge.exe\" start"
            )
        except requests.Timeout:
            raise WebBridgeError(f"WebBridge 命令超时: {action}")

        if not result.get("ok"):
            err = result.get("error", {}).get("message", str(result))
            raise WebBridgeError(f"{action} 失败: {err}")

        return result.get("data", result)

    # ── 公开 API ─────────────────────────────────────────────────────

    def navigate(self, url: str, new_tab: bool = False,
                 group_title: str | None = None) -> dict:
        """导航到 URL。首次调用自动打开新 tab 并设置分组标题。"""
        args: dict[str, Any] = {"url": url, "newTab": new_tab}
        if group_title:
            args["group_title"] = group_title
        return self._post("navigate", args)

    def snapshot(self) -> dict:
        """返回页面无障碍树（含 @e 引用）。"""
        return self._post("snapshot")

    def click(self, selector: str) -> dict:
        """点击元素。selector 支持 CSS 选择器或 @e 引用。"""
        return self._post("click", {"selector": selector})

    def fill(self, selector: str, value: str) -> dict:
        """填充输入框（清除后填入）。支持 <input> / <textarea> / contenteditable。"""
        return self._post("fill", {"selector": selector, "value": value})

    def evaluate(self, code: str) -> Any:
        """执行 JavaScript 代码，返回结果。"""
        result = self._post("evaluate", {"code": code})
        value = result.get("value")
        # 尝试解析 JSON
        if isinstance(value, str) and value.strip().startswith(("{", "[")):
            try:
                return json.loads(value)
            except json.JSONDecodeError:
                pass
        return value

    def screenshot(self, path: str, format: str = "png",
                   quality: int = 80) -> dict:
        """截图保存到 path。返回 {format, path, sizeBytes}。"""
        return self._post("screenshot", {
            "path": path,
            "format": format,
            "quality": quality,
        })

    def cdp(self, method: str, params: dict | None = None) -> dict:
        """发送原始 CDP (Chrome DevTools Protocol) 命令。"""
        return self._post("cdp", {"method": method, "params": params or {}})

    def find_tab(self, url: str, active: bool = False) -> dict:
        """切换到已打开的 tab（按 URL 匹配）。"""
        return self._post("find_tab", {"url": url, "active": active})

    def list_tabs(self) -> list[dict]:
        """列出当前 session 的所有 tab。"""
        result = self._post("list_tabs")
        return result.get("tabs", [])

    def close_session(self) -> dict:
        """关闭当前 session 的所有 tab。"""
        return self._post("close_session")

    # ── 高级工具方法 ─────────────────────────────────────────────────

    def wait_for(self, css_selector: str, timeout: float = 15) -> bool:
        """轮询等待 CSS 选择器匹配的元素出现在 DOM 中。"""
        deadline = time.time() + timeout
        interval = 0.5
        while time.time() < deadline:
            count = self.evaluate(
                f"document.querySelectorAll({json.dumps(css_selector)}).length"
            )
            if count > 0:
                return True
            time.sleep(interval)
            interval = min(interval * 1.3, 2.0)
        return False

    def wait_text_in(self, css_selector: str, text: str,
                     timeout: float = 15) -> bool:
        """等待某个 CSS 选择器下的文本变化（包含指定文本）。"""
        deadline = time.time() + timeout
        interval = 0.5
        while time.time() < deadline:
            current = self.evaluate(
                f"(document.querySelector({json.dumps(css_selector)}) || {{}}).textContent || ''"
            )
            if text in str(current):
                return True
            time.sleep(interval)
            interval = min(interval * 1.3, 2.0)
        return False

    def element_exists(self, css_selector: str) -> bool:
        """检查元素是否存在于 DOM 中。"""
        count = self.evaluate(
            f"document.querySelectorAll({json.dumps(css_selector)}).length"
        )
        return count > 0
