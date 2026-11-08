"""生产计划局域网共享：本机计算完成后通过 HTTP 提供快照，其他主机可直接取用。

工作方式：
- 每台机器计算完 5 天计划后，把结果以 JSON 快照发布到本机 8790 端口；
- 其他机器刷新时先查询设置中的 3 个终端 IP，取时间戳最新且日期匹配的快照，
  直接使用而不再读取 Excel 重新计算；取不到时才本机计算。
"""

from __future__ import annotations

import json
import threading
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from app.utils.debug_log import info, warn

PLAN_SHARE_PORT = 8790


class _ShareHandler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:  # noqa: N802
        path = self.path.split("?", 1)[0].rstrip("/") or "/"
        if path == "/":
            snapshot = self.server.snapshot_provider()  # type: ignore[attr-defined]
            self._respond(
                {
                    "ok": True,
                    "service": "plan-share",
                    "has_data": snapshot is not None,
                    "synced_at": (snapshot or {}).get("synced_at", ""),
                    "today": (snapshot or {}).get("today", ""),
                    "usage": "GET /plan 获取计划快照",
                }
            )
            return
        if path == "/plan":
            snapshot = self.server.snapshot_provider()  # type: ignore[attr-defined]
            if snapshot is None:
                self._respond(
                    {
                        "ok": False,
                        "error": "暂无数据：本机尚未完成一次计划计算，"
                        "请先登录应用并等待首页数据同步完成",
                    },
                    status=404,
                )
            else:
                self._respond(snapshot)
            return
        self._respond({"ok": False, "error": "未知接口"}, status=404)

    def log_message(self, *args) -> None:  # 静默访问日志
        return

    def _respond(self, payload: dict, status: int = 200) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        try:
            self.wfile.write(body)
        except OSError:
            pass


class PlanShareService:
    """本机计划快照的发布者。"""

    def __init__(self) -> None:
        self._snapshot: dict | None = None
        self._lock = threading.Lock()
        self._server: ThreadingHTTPServer | None = None
        self._start_server()

    def _start_server(self) -> None:
        try:
            server = ThreadingHTTPServer(("0.0.0.0", PLAN_SHARE_PORT), _ShareHandler)
        except OSError as exc:
            warn("[计划共享]", f"端口 {PLAN_SHARE_PORT} 启动失败：{exc}（本机仅作为获取方）")
            return
        server.snapshot_provider = self.snapshot  # type: ignore[attr-defined]
        server.daemon_threads = True
        self._server = server
        thread = threading.Thread(
            target=server.serve_forever, kwargs={"poll_interval": 0.5}, daemon=True
        )
        thread.start()
        info("[计划共享]", f"已在 0.0.0.0:{PLAN_SHARE_PORT} 提供计划快照接口")

    def publish(self, snapshot: dict) -> None:
        with self._lock:
            self._snapshot = snapshot

    def snapshot(self) -> dict | None:
        with self._lock:
            return self._snapshot

    def stop(self) -> None:
        if self._server is not None:
            self._server.shutdown()
            self._server.server_close()
            self._server = None


def _fetch_one(ip: str, today: date, port: int, timeout: float) -> dict | None:
    url = f"http://{ip}:{port}/plan"
    try:
        with urllib.request.urlopen(url, timeout=timeout) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except Exception as exc:  # noqa: BLE001
        info("[计划共享]", f"{ip} 无可用数据：{type(exc).__name__}")
        return None
    if not isinstance(payload, dict) or payload.get("today") != today.isoformat():
        return None
    rows = payload.get("rows")
    if not isinstance(rows, dict):
        return None
    total = sum(
        len(items)
        for mode_rows in rows.values()
        if isinstance(mode_rows, dict)
        for items in mode_rows.values()
        if isinstance(items, list)
    )
    if total == 0:
        return None
    return payload


def fetch_latest_remote(
    ips: list[str] | tuple[str, ...],
    today: date,
    port: int = PLAN_SHARE_PORT,
    timeout: float = 1.5,
) -> tuple[dict, str] | None:
    """并发查询各终端 IP 的计划快照，返回 (最新快照, 来源IP)；无可用数据返回 None。

    只接受 today 与本机一致的快照，避免拿到其他日期的旧数据。
    """
    targets = [ip.strip() for ip in ips if ip and ip.strip()]
    if not targets:
        return None
    best: tuple[dict, str] | None = None
    with ThreadPoolExecutor(max_workers=len(targets)) as pool:
        futures = {
            pool.submit(_fetch_one, ip, today, port, timeout): ip for ip in targets
        }
        for future in as_completed(futures):
            payload = future.result()
            if payload is None:
                continue
            ip = futures[future]
            synced_at = str(payload.get("synced_at", ""))
            if best is None or synced_at > str(best[0].get("synced_at", "")):
                best = (payload, ip)
    return best
