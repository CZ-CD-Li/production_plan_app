"""产线代号与角色权限。"""

from __future__ import annotations

import re

ASSEMBLY_LINES: tuple[str, ...] = tuple(f"C{i}" for i in range(1, 24))
WELDING_LINES: tuple[str, ...] = tuple(f"B{i}" for i in range(1, 12))
MACHINING_LINES: tuple[str, ...] = (
    tuple(f"A{i}" for i in range(1, 4)) + WELDING_LINES
)
ALL_LINES: tuple[str, ...] = (
    tuple(f"A{i}" for i in range(1, 4))
    + WELDING_LINES
    + ASSEMBLY_LINES
)

_CODE_PATTERN = re.compile(r"[ABC]\d+", re.IGNORECASE)


def _line_key(line: str) -> tuple[str, int]:
    return line[0], int(line[1:])


def allowed_lines(role: str) -> tuple[str, ...]:
    """按角色返回可见产线；角色不含装配/焊接/机加时返回全部产线。"""
    text = (role or "").strip()
    lines: list[str] = []
    if "装配" in text:
        lines.extend(ASSEMBLY_LINES)
    if "焊接" in text:
        lines.extend(WELDING_LINES)
    if "机加" in text:
        lines.extend(MACHINING_LINES)
    if not lines:
        return ALL_LINES
    return tuple(sorted(dict.fromkeys(lines), key=_line_key))


def normalise_line(value: str) -> str | None:
    """把输入规范为合法产线代号（大小写不敏感），不匹配返回 None。"""
    text = str(value or "").strip().upper()
    if not text:
        return None
    return text if text in ALL_LINES else None


def line_codes_in_text(value: str) -> set[str]:
    """提取文本中的全部产线代号（如“B7/C12E12” → {B7, C12, E12}）。"""
    return {match.upper() for match in _CODE_PATTERN.findall(str(value or ""))}


def line_in_text(value: str, lines: tuple[str, ...]) -> str | None:
    """从“B7/C12E12”这类组合文本中取第一个在允许范围内的产线代号。"""
    allowed = {line.upper() for line in lines}
    for match in _CODE_PATTERN.findall(str(value or "")):
        code = match.upper()
        if code in allowed:
            return code
    return None
