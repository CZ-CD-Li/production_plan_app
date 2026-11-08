"""后减作业指导书生成：从 PDF 图纸生成初版作业指导书。"""

from __future__ import annotations

import re
import shutil
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import load_workbook

from app.config import DATA_DIR, ROOT_DIR
from app.table_ocr.recognizer import TableRow, recognize_pdf

LINE_OPTIONS = tuple(f"C{index}" for index in range(1, 24))
DEFAULT_LINE = "C9"
TEMPLATE_DIR = DATA_DIR / "templates"
OUTPUT_DIR = DATA_DIR / "作业指导书"
TEMPLATE_SUFFIX = "标准作业指导书模板.xlsx"

POSITION_PAD = "定位垫"
POSITION_PAD_SHEET = "拉装定位垫"


class InstructionError(Exception):
    pass


@dataclass
class Drawing:
    path: str
    code: str = ""
    display_name: str = ""
    rows: list[TableRow] = field(default_factory=list)
    problems: list[str] = field(default_factory=list)

    @property
    def is_assembly(self) -> bool:
        return "总成" in self.display_name


@dataclass
class InstructionResult:
    output: Path
    product_code: str
    line: str
    drawings: list[Drawing] = field(default_factory=list)
    removed_sheets: list[str] = field(default_factory=list)
    changes: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def template_path(line: str) -> Path:
    """模板优先取数据目录，缺失时从打包资源复制一份（与 SPC 模板同一套路）。"""
    filename = f"{line}{TEMPLATE_SUFFIX}"
    cached = TEMPLATE_DIR / filename
    if cached.is_file():
        return cached
    candidates: list[Path] = []
    meipass = getattr(sys, "_MEIPASS", "")
    if meipass:
        candidates.append(Path(meipass) / "data" / "templates" / filename)
        candidates.append(Path(meipass) / filename)
    if getattr(sys, "frozen", False):
        candidates.append(Path(sys.executable).resolve().parent / filename)
    candidates.append(ROOT_DIR / "data" / "templates" / filename)
    for source in candidates:
        if source.is_file():
            TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, cached)
            return cached
    raise InstructionError(
        f"未找到 {line} 的标准作业指导书模板（{filename}），"
        f"请把模板放到 {TEMPLATE_DIR} 下。"
    )


def _split_name(stem: str) -> tuple[str, str]:
    """把文件名拆成代号 + 名称，如 3C027060活塞杆组件 -> 3C027060 / 活塞杆组件。"""
    match = re.match(r"[0-9A-Za-z]+", stem)
    if not match:
        return "", stem
    return match.group(0), stem[match.end() :].strip(" -_")


def load_drawings(paths: list[str]) -> list[Drawing]:
    drawings: list[Drawing] = []
    for path in paths:
        item = Path(path)
        if item.suffix.lower() != ".pdf":
            raise InstructionError(f"仅支持 PDF 图纸：{item.name}")
        code, display = _split_name(item.stem)
        drawing = Drawing(path=str(item), code=code, display_name=display or item.stem)
        try:
            recognition = recognize_pdf(str(item))
        except Exception as exc:  # noqa: BLE001
            drawing.problems.append(f"识别失败：{exc}")
        else:
            drawing.rows = recognition.rows
            if recognition.warnings:
                drawing.problems.extend(recognition.warnings)
            if not drawing.rows:
                drawing.problems.append("未识别到明细表")
        drawings.append(drawing)
    return drawings


def _read_page_text(path: str) -> str:
    """图纸没有表格结构时，用整页 OCR 找关键字。"""
    import pymupdf

    from app.table_ocr.recognizer import _get_ocr_engine, _ocr_image, _render_crop

    engine = _get_ocr_engine()
    chunks: list[str] = []
    doc = pymupdf.open(path)
    try:
        for page in doc:
            text = page.get_text().strip()
            if text:
                chunks.append(text)
                continue
            joined = ""
            for angle in (270, 90, 0):
                items = _ocr_image(_render_crop(page, page.rect, 110, angle), engine)
                joined = " ".join(item.text for item in items)
                if POSITION_PAD in joined or "技术要求" in joined:
                    break
            chunks.append(joined)
    finally:
        doc.close()
    return "\n".join(chunks)


def _looks_like_bom(rows: list[TableRow]) -> bool:
    """判断识别结果是否真的是明细表，标题栏文字不算。"""
    real = [
        row
        for row in rows
        if re.fullmatch(r"[0-9A-Za-z][0-9A-Za-z\-./]{4,}", (row.code or "").strip())
        and (row.name or "").strip()
    ]
    return len(real) >= 3


def has_position_pad(drawing: Drawing) -> bool | None:
    """判断阻尼器图纸里是否有定位垫。

    只有明确识别出明细表且其中没有该零件时才返回 False；
    True 表示确认存在，None 表示无法判断（调用方应保留工序）。
    """
    if any(POSITION_PAD in row.name for row in drawing.rows):
        return True
    if _looks_like_bom(drawing.rows):
        return False
    try:
        text = _read_page_text(drawing.path)
    except Exception:  # noqa: BLE001
        return None
    if POSITION_PAD in text:
        return True
    return None


# ---------- 单元格规划 ----------


def _card_edits(ws, rows_by_code: dict[str, list[TableRow]], edits: list) -> None:
    """工艺卡片图：按图纸明细逐行填 代号/名称/数量，多出的模板行清空。

    模板首行与末行的名称是范围锚点（如“支承垫…支承垫”），只取图纸中对应的那一段。
    """
    header_row = None
    for row in range(1, min(ws.max_row, 15) + 1):
        if str(ws.cell(row=row, column=4).value or "").strip() == "序号":
            header_row = row
            break
    if header_row is None:
        return
    code = str(ws.cell(row=6, column=4).value or "").strip().upper()
    rows = rows_by_code.get(code)
    if not rows:
        return

    start = header_row + 1
    capacity = 0
    while start + capacity <= ws.max_row and ws.cell(
        row=start + capacity, column=5
    ).value not in (None, ""):
        capacity += 1
    if capacity == 0:
        return

    first_name = str(ws.cell(row=start, column=6).value or "").strip()
    last_name = str(ws.cell(row=start + capacity - 1, column=6).value or "").strip()
    picked = _slice_by_anchor(rows, first_name, last_name, capacity)

    for offset in range(capacity):
        row = start + offset
        if offset < len(picked):
            item = picked[offset]
            edits.append((row, 5, item.code))
            edits.append((row, 6, item.name))
            edits.append((row, 7, item.quantity))
        else:
            edits.append((row, 5, ""))
            edits.append((row, 6, ""))
            edits.append((row, 7, ""))


def _slice_by_anchor(
    rows: list[TableRow], first_name: str, last_name: str, capacity: int
) -> list[TableRow]:
    """按模板首末名称在图纸明细里截取对应的连续区间。"""

    def find(name: str, start: int = 0) -> int | None:
        for index in range(start, len(rows)):
            if rows[index].name.strip() == name:
                return index
        return None

    if first_name and last_name:
        start = find(first_name)
        if start is not None:
            end = find(last_name, start + 1)
            if end is not None:
                return rows[start : end + 1]
    return rows[:capacity]


_BOM_KEYS = ("名称", "代号", "数量")


def _bom_edits(ws, parts: dict[str, TableRow], edits: list) -> None:
    """部件构成一览表：按代号校正名称与数量（版次不覆盖，留待人工确认）。"""
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=8):
        head = row[0].value
        if not isinstance(head, str) or "部件构成一览表" not in head:
            continue
        header_row = row[0].row + 1
        columns: dict[str, int] = {}
        for column in range(1, min(ws.max_column, 10) + 1):
            text = str(ws.cell(row=header_row, column=column).value or "").strip()
            if text in _BOM_KEYS:
                columns.setdefault(text, column)
        name_col = columns.get("名称")
        code_col = columns.get("代号")
        qty_col = columns.get("数量")
        if not name_col or not code_col:
            return

        cursor = header_row + 1
        while cursor <= ws.max_row:
            name = ws.cell(row=cursor, column=name_col).value
            code = ws.cell(row=cursor, column=code_col).value
            if isinstance(name, str) and ("特记事项" in name or "作业前" in name):
                break
            if code in (None, "") and name in (None, ""):
                break
            part = parts.get(str(code or "").strip().upper())
            if part is not None:
                if part.name and str(name or "").strip() != part.name:
                    edits.append((cursor, name_col, part.name))
                if qty_col and part.quantity:
                    current = str(
                        ws.cell(row=cursor, column=qty_col).value or ""
                    ).strip()
                    if current != part.quantity:
                        edits.append((cursor, qty_col, part.quantity))
            cursor += 1
        return


# ---------- XLSX 直接改 XML，保留嵌入图纸 ----------


def _escape(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _set_cell(xml: str, ref: str, value: str) -> tuple[str, bool]:
    pattern = re.compile(
        rf'<c r="{ref}"(?:[^>]*)>(?:[^<]*<v>[^<]*</v>)?</c>|<c r="{ref}"[^>]*/>'
    )
    match = pattern.search(xml)
    if not match:
        return xml, False
    style = re.search(r's="(\d+)"', match.group(0))
    attr = f' s="{style.group(1)}"' if style else ""
    text = str(value)
    if text == "":
        body = ""
    elif text.isdigit():
        body = f"<v>{text}</v>"
    else:
        body = f'<is><t xml:space="preserve">{_escape(text)}</t></is>'
        attr += ' t="inlineStr"'
    replacement = f'<c r="{ref}"{attr}>{body}</c>'
    return xml[: match.start()] + replacement + xml[match.end() :], True


def _col_letter(column: int) -> str:
    letters = ""
    while column > 0:
        column, remainder = divmod(column - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _sheet_parts(data: dict[str, bytes]) -> dict[str, tuple[str, str]]:
    workbook = data["xl/workbook.xml"].decode("utf-8")
    rels = data["xl/_rels/workbook.xml.rels"].decode("utf-8")
    targets = dict(
        re.findall(r'<Relationship[^>]*Id="(rId\d+)"[^>]*Target="([^"]+)"', rels)
    )
    mapping: dict[str, tuple[str, str]] = {}
    for name, rid in re.findall(
        r'<sheet[^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"', workbook
    ):
        target = targets.get(rid, "").lstrip("/")
        part = target if target.startswith("xl/") else f"xl/{target}"
        mapping[name] = (rid, part)
    return mapping


def _drop_sheet(data: dict[str, bytes], name: str, rid: str, part: str) -> None:
    workbook = data["xl/workbook.xml"].decode("utf-8")
    workbook = re.sub(rf'<sheet[^>]*name="{re.escape(name)}"[^>]*/>', "", workbook)
    # 删表后按序号索引的旧名称会错位，这些名称都已失效，整段去掉
    workbook = re.sub(r"<definedNames>.*?</definedNames>", "", workbook, flags=re.S)
    data["xl/workbook.xml"] = workbook.encode("utf-8")

    rels = data["xl/_rels/workbook.xml.rels"].decode("utf-8")
    rels = re.sub(rf'<Relationship[^>]*Id="{rid}"[^>]*/>', "", rels)
    data["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")

    data.pop(part, None)
    data.pop(part.replace("worksheets/", "worksheets/_rels/") + ".rels", None)

    content_types = data["[Content_Types].xml"].decode("utf-8")
    content_types = re.sub(
        rf'<Override[^>]*PartName="/{re.escape(part)}"[^>]*/>', "", content_types
    )
    data["[Content_Types].xml"] = content_types.encode("utf-8")


def generate(
    pdf_paths: list[str],
    line: str = DEFAULT_LINE,
    out_dir: Path | str | None = None,
    template: Path | str | None = None,
) -> InstructionResult:
    if not pdf_paths:
        raise InstructionError("请先导入至少一份 PDF 图纸。")
    template_file = Path(template) if template else template_path(line)
    if not template_file.is_file():
        raise InstructionError(f"模板文件不存在：{template_file}")

    drawings = load_drawings(pdf_paths)
    usable = [item for item in drawings if item.rows]
    if not usable:
        raise InstructionError("所有图纸都没有识别出明细表，无法生成作业指导书。")

    product = next((d for d in drawings if d.is_assembly), usable[0])
    product_code = product.code
    if not product_code:
        raise InstructionError(f"无法从图纸文件名解析产品代号：{Path(product.path).name}")

    rows_by_code = {d.code.upper(): d.rows for d in usable}
    parts: dict[str, TableRow] = {}
    for drawing in usable:
        for row in drawing.rows:
            if row.code:
                parts.setdefault(row.code.upper(), row)

    target_dir = Path(out_dir) if out_dir else OUTPUT_DIR
    target_dir.mkdir(parents=True, exist_ok=True)
    output = target_dir / f"{product_code}-作业指导书（初版）.xlsx"
    shutil.copy2(template_file, output)

    with zipfile.ZipFile(output) as zf:
        entries = [(info, zf.read(info.filename)) for info in zf.infolist()]
    data = {info.filename: content for info, content in entries}
    sheets = _sheet_parts(data)

    workbook = load_workbook(output)
    edits: dict[str, list[tuple[int, int, str]]] = {}
    removed: list[str] = []
    warnings: list[str] = []

    damper = next((d for d in drawings if "阻尼器" in d.display_name), None)
    delete_pad = False
    if damper is None:
        warnings.append("未导入阻尼器图纸，无法判断是否需要拉装定位垫工序，已保留该工序。")
    else:
        detected = has_position_pad(damper)
        if detected is None:
            warnings.append("阻尼器图纸未能确认定位垫零件，已保留拉装定位垫工序，请人工确认。")
        else:
            delete_pad = not detected

    for name, (rid, part) in list(sheets.items()):
        if delete_pad and POSITION_PAD_SHEET in name:
            _drop_sheet(data, name, rid, part)
            removed.append(name)
            continue
        if name not in workbook.sheetnames:
            continue
        worksheet = workbook[name]
        bucket: list[tuple[int, int, str]] = []
        if "工艺卡片图" in name:
            _card_edits(worksheet, rows_by_code, bucket)
        _bom_edits(worksheet, parts, bucket)
        if bucket:
            edits[name] = bucket
    workbook.close()

    changes: list[str] = []
    for name, bucket in edits.items():
        filename = sheets[name][1]
        xml = data[filename].decode("utf-8")
        applied = 0
        card_rows = 0
        for row, column, value in bucket:
            ref = f"{_col_letter(column)}{row}"
            xml, ok = _set_cell(xml, ref, value)
            if ok:
                applied += 1
                if "工艺卡片图" in name and column == 5:
                    card_rows += 1
        data[filename] = xml.encode("utf-8")
        if card_rows:
            changes.append(f"{name}：更新 {card_rows} 行明细")
        elif applied:
            changes.append(f"{name}：校正 {applied} 处名称/数量")

    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zf:
        for info, _ in entries:
            if info.filename in data:
                zf.writestr(info.filename, data[info.filename])

    if delete_pad:
        changes.append("阻尼器图纸无定位垫，已删除拉装定位垫工序")
    elif damper is not None:
        changes.append("阻尼器图纸含定位垫，已保留拉装定位垫工序")

    for drawing in drawings:
        for problem in drawing.problems:
            warnings.append(f"{Path(drawing.path).name}：{problem}")

    return InstructionResult(
        output=output,
        product_code=product_code,
        line=line,
        drawings=drawings,
        removed_sheets=removed,
        changes=changes,
        warnings=warnings,
    )
