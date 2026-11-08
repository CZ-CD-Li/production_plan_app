"""诊断日志：同时输出到控制台与 data/debug.log。"""

from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path
from threading import Lock

from app.config import DATA_DIR

LOG_PATH = DATA_DIR / "debug.log"

_lock = Lock()
_file_ready = False


def _ensure_file() -> None:
    global _file_ready
    if _file_ready:
        return
    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    _file_ready = True


def _write_file(text: str) -> None:
    try:
        _ensure_file()
        with LOG_PATH.open("a", encoding="utf-8") as handle:
            handle.write(text + "\n")
    except OSError:
        pass


def start_session(title: str) -> None:
    """开启一次新的诊断会话，在控制台与日志文件中写入分隔头。"""
    now = datetime.now()
    header = (
        "\n"
        + "=" * 78
        + f"\n[{now:%Y-%m-%d %H:%M:%S}] {title}\n"
        + f"诊断日志文件：{LOG_PATH}\n"
        + "=" * 78
    )
    with _lock:
        try:
            print(header, flush=True)
        except Exception:  # noqa: BLE001
            pass
        _write_file(header)


def log(tag: str, message: str = "", level: str = "INFO") -> None:
    """输出一条带标签的诊断日志。"""
    stamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
    line = f"{stamp} [{level:<5}] {tag}"
    if message:
        line = f"{line} {message}"
    with _lock:
        try:
            print(line, flush=True)
        except Exception:  # noqa: BLE001
            # 控制台编码异常时不影响日志文件记录
            try:
                sys.stdout.buffer.write((line + "\n").encode("utf-8", "replace"))
                sys.stdout.buffer.flush()
            except Exception:  # noqa: BLE001
                pass
        _write_file(line)


def info(tag: str, message: str = "") -> None:
    log(tag, message, "INFO")


def warn(tag: str, message: str = "") -> None:
    log(tag, message, "WARN")


def error(tag: str, message: str = "") -> None:
    log(tag, message, "ERROR")


def detail(tag: str, message: str = "") -> None:
    log(tag, message, "DEBUG")


def exception(tag: str, exc: BaseException) -> None:
    log(tag, f"{type(exc).__name__}: {exc}", "ERROR")
