"""由程序内图标生成 Windows 可执行文件图标（build_tools/app.ico）。"""

from __future__ import annotations

import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PyQt5.QtWidgets import QApplication  # noqa: E402
from PyQt5.QtCore import QBuffer, QByteArray, QIODevice, QSize  # noqa: E402

from app.utils.icons import app_icon  # noqa: E402


def main() -> int:
    app = QApplication.instance() or QApplication(sys.argv)
    icon = app_icon()
    target = Path(__file__).resolve().parent / "app.ico"

    payload = []
    for size in (16, 24, 32, 48, 64, 128, 256):
        image = icon.pixmap(QSize(size, size)).toImage()
        if image.isNull():
            continue
        # ICO 内的每个尺寸都用 PNG 编码，Windows Vista 及以上均支持
        array = QByteArray()
        buffer = QBuffer(array)
        buffer.open(QIODevice.WriteOnly)
        image.save(buffer, "PNG")
        buffer.close()
        payload.append((size, bytes(array)))

    if not payload:
        print("图标生成失败：没有可用的位图")
        return 1

    import struct

    header = struct.pack("<HHH", 0, 1, len(payload))
    offset = 6 + 16 * len(payload)
    entries = b""
    blob = b""
    for size, data in payload:
        width = 0 if size >= 256 else size
        entries += struct.pack(
            "<BBBBHHII", width, width, 0, 0, 1, 32, len(data), offset
        )
        blob += data
        offset += len(data)

    target.write_bytes(header + entries + blob)
    print(f"已生成 {target}（{target.stat().st_size} 字节）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
