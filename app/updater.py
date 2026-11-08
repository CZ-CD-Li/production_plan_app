"""启动时检查局域网共享目录中的新版本安装程序，并强制升级。"""

from __future__ import annotations

import os
import re
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path

from PyQt5.QtWidgets import QMessageBox

from app.config import (
    APP_NAME,
    APP_VERSION,
    ROOT_DIR,
    UPDATE_SHARE,
    UPDATE_SHARE_ENV,
)
from app.utils.debug_log import error, info

INSTALLER_RE = re.compile(
    r"^" + re.escape(APP_NAME) + r"_安装程序_v(\d+(?:\.\d+){1,2})\.exe$",
    re.IGNORECASE,
)


@dataclass
class UpdateInfo:
    version: tuple[int, ...]
    version_text: str
    path: Path


def _version_tuple(text: str) -> tuple[int, ...]:
    parts = [int(piece) for piece in re.findall(r"\d+", text)[:3]]
    while len(parts) < 3:
        parts.append(0)
    return tuple(parts)


def share_path() -> Path:
    override = os.environ.get(UPDATE_SHARE_ENV, "").strip()
    return Path(override) if override else Path(UPDATE_SHARE)


def find_update(share: Path | None = None) -> UpdateInfo | None:
    """返回共享目录中版本高于当前程序的最新安装包，找不到返回 None。"""
    folder = share or share_path()
    try:
        if not folder.is_dir():
            info("[自动更新]", f"共享目录不可访问，跳过检查：{folder}")
            return None
        entries = list(folder.iterdir())
    except OSError as exc:
        info("[自动更新]", f"共享目录不可访问，跳过检查：{folder}（{exc}）")
        return None

    current = _version_tuple(APP_VERSION)
    best: UpdateInfo | None = None
    for entry in entries:
        if not entry.is_file():
            continue
        match = INSTALLER_RE.match(entry.name)
        if match is None:
            continue
        version = _version_tuple(match.group(1))
        if version <= current:
            continue
        if best is None or version > best.version:
            best = UpdateInfo(
                version=version,
                version_text=".".join(str(piece) for piece in version),
                path=entry,
            )
    return best


def install_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return ROOT_DIR


def _write_script(info: UpdateInfo, target: Path) -> Path:
    """生成升级批处理：等旧程序退出 → 静默卸载 → 同目录静默安装 → 启动新版本。"""
    script = Path(tempfile.gettempdir()) / f"{APP_NAME}_update.cmd"
    uninstaller = target / "unins000.exe"
    exe = target / f"{APP_NAME}.exe"
    lines = [
        "@echo off",
        "timeout /t 3 /nobreak >nul",
    ]
    if uninstaller.is_file():
        lines.append(
            f'start /wait "" "{uninstaller}" /VERYSILENT /SUPPRESSMSGBOXES /NORESTART'
        )
    lines.append(
        f'start /wait "" "{info.path}" /VERYSILENT /SUPPRESSMSGBOXES '
        f'/NORESTART /DIR="{target}"'
    )
    lines.append(f'start "" "{exe}"')
    lines.append('del "%~f0"')
    script.write_text("\r\n".join(lines) + "\r\n", encoding="gbk")
    return script


def run_installer(info: UpdateInfo, target: Path) -> None:
    """启动升级脚本，调用方随后立即退出程序。"""
    script = _write_script(info, target)
    flags = 0
    for name in ("DETACHED_PROCESS", "CREATE_NEW_PROCESS_GROUP", "CREATE_NO_WINDOW"):
        flags |= getattr(subprocess, name, 0)
    subprocess.Popen(  # noqa: S603
        ["cmd.exe", "/c", str(script)],
        close_fds=True,
        creationflags=flags,
    )


def check_and_prompt() -> bool:
    """检查更新；返回 True 表示可继续使用当前版本。"""
    if not getattr(sys, "frozen", False):
        return True
    if "--skip-update" in sys.argv:
        return True

    info = find_update()
    if info is None:
        return True

    info("[自动更新]", f"发现新版本 {info.version_text}：{info.path}")
    box = QMessageBox()
    box.setWindowTitle("发现新版本")
    box.setIcon(QMessageBox.Information)
    box.setText(f"新版本 {info.version_text} 已就绪，请安装新版本。")
    box.setInformativeText(
        f"当前版本：{APP_VERSION}\n最新版本：{info.version_text}\n\n"
        "安装将自动卸载旧版本并安装到原目录，安装完成后自动启动。\n"
        "不支持继续使用旧版本。"
    )
    install_button = box.addButton("立即安装", QMessageBox.AcceptRole)
    exit_button = box.addButton("退出程序", QMessageBox.RejectRole)
    box.setDefaultButton(install_button)
    box.exec_()

    if box.clickedButton() is not install_button:
        QMessageBox.information(
            None, "必须升级", "旧版本已停用，请安装新版本后再使用。"
        )
        return False

    try:
        target = install_dir()
        info("[自动更新]", f"开始升级到 {info.version_text}，安装目录={target}")
        run_installer(info, target)
    except OSError as exc:
        error("[自动更新]", f"启动安装程序失败：{type(exc).__name__}: {exc}")
        QMessageBox.critical(
            None,
            "升级失败",
            f"无法启动安装程序：{exc}\n\n请手动运行：\n{info.path}",
        )
    return False
