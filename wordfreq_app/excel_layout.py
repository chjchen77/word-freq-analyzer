"""Automatic worksheet and header-row detection for spreadsheet inputs."""

from __future__ import annotations

import os
import posixpath
import zipfile
from functools import lru_cache
from pathlib import Path
from typing import Any, Iterable
from xml.etree import ElementTree as ET

HEADER_SCAN_ROWS = 20
HEADER_HINTS = (
    "公司", "企业", "股票", "证券", "代码", "编号", "日期", "时间", "年份",
    "year", "date", "code", "company", "文本", "内容", "正文", "回复",
    "问题", "标题", "摘要", "text", "content", "reply", "question", "body",
)


def _clean(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return "" if text.lower() in {"nan", "none", "nat"} else text


def choose_header_row(rows: Iterable[Iterable[Any]]) -> tuple[int, float]:
    """Choose a zero-based header row and return a simple 0–1 confidence."""
    prepared = [list(row) for row in rows]
    if not prepared:
        return 0, 0.0

    scored: list[tuple[float, int]] = []
    max_width = max((len(row) for row in prepared), default=1)
    for index, row in enumerate(prepared[:HEADER_SCAN_ROWS]):
        values = [_clean(value) for value in row]
        nonempty = [value for value in values if value]
        if not nonempty:
            scored.append((-1000.0, index))
            continue
        unique_ratio = len(set(nonempty)) / len(nonempty)
        string_like = sum(not value.replace(".", "", 1).isdigit() for value in nonempty)
        hints = sum(
            1 for value in nonempty
            if any(hint in value.lower() for hint in HEADER_HINTS)
        )
        following_nonempty = 0
        if index + 1 < len(prepared):
            following_nonempty = sum(bool(_clean(v)) for v in prepared[index + 1])
        # A real header tends to be wide, unique, textual, followed by data, and
        # close to the top. Domain hints deliberately receive the largest weight.
        score = (
            len(nonempty) * 3.0
            + unique_ratio * 4.0
            + string_like * 1.5
            + hints * 15.0
            + min(following_nonempty, len(nonempty)) * 0.5
            + max(0, 5 - index) * 0.75
        )
        if len(nonempty) == 1:
            score -= 12.0
        scored.append((score, index))

    scored.sort(key=lambda item: (-item[0], item[1]))
    best_score, best_index = scored[0]
    theoretical = max(20.0, max_width * 6.0 + 15.0)
    confidence = max(0.0, min(1.0, best_score / theoretical))
    return best_index, round(confidence, 3)


_SHARED_STRING_REF = object()


def _xml_name(value: str) -> str:
    """Return an XML tag/attribute name without its namespace."""
    return value.rsplit("}", 1)[-1]


def _cell_column_index(reference: str) -> int:
    """Return a zero-based column index from an XLSX cell reference."""
    result = 0
    for character in reference:
        if not character.isalpha():
            break
        result = result * 26 + (ord(character.upper()) - ord("A") + 1)
    return max(0, result - 1)


def _cell_value(cell: ET.Element) -> Any:
    """Extract a value from a worksheet cell without loading openpyxl."""
    cell_type = cell.attrib.get("t", "")
    if cell_type == "inlineStr":
        return "".join(
            part.text or "" for part in cell.iter() if _xml_name(part.tag) == "t"
        )

    raw_value = ""
    for child in cell:
        if _xml_name(child.tag) == "v":
            raw_value = child.text or ""
            break
    if cell_type == "s":
        try:
            return (_SHARED_STRING_REF, int(raw_value))
        except ValueError:
            return None
    return raw_value or None


def _read_xlsx_rows(archive: zipfile.ZipFile, sheet_path: str) -> list[list[Any]]:
    """Read only the first header-scan rows from a worksheet XML stream."""
    rows: dict[int, list[Any]] = {}
    max_row = HEADER_SCAN_ROWS + 1
    with archive.open(sheet_path) as stream:
        for _event, element in ET.iterparse(stream, events=("end",)):
            if _xml_name(element.tag) != "row":
                continue
            try:
                row_number = int(element.attrib.get("r", "0"))
            except ValueError:
                row_number = 0
            if row_number <= 0:
                row_number = len(rows) + 1
            if row_number > max_row:
                element.clear()
                break

            values_by_column: dict[int, Any] = {}
            for cell in element:
                if _xml_name(cell.tag) != "c":
                    continue
                column_index = _cell_column_index(cell.attrib.get("r", "A1"))
                values_by_column[column_index] = _cell_value(cell)
            if values_by_column:
                values = [None] * (max(values_by_column) + 1)
                for column_index, value in values_by_column.items():
                    values[column_index] = value
                rows[row_number] = values
            element.clear()

    return [rows.get(index, []) for index in range(1, max_row + 1)]


def _resolve_shared_strings(
    archive: zipfile.ZipFile, rows: list[list[Any]],
) -> list[list[Any]]:
    """Resolve only the shared-string indexes present in the scanned rows."""
    needed = {
        value[1]
        for row in rows
        for value in row
        if isinstance(value, tuple) and len(value) == 2 and value[0] is _SHARED_STRING_REF
    }
    if not needed:
        return rows

    try:
        stream = archive.open("xl/sharedStrings.xml")
    except KeyError:
        return rows

    resolved: dict[int, str] = {}
    max_needed = max(needed)
    with stream:
        index = 0
        for _event, element in ET.iterparse(stream, events=("end",)):
            if _xml_name(element.tag) != "si":
                continue
            if index in needed:
                resolved[index] = "".join(element.itertext())
            index += 1
            element.clear()
            if index > max_needed and needed.issubset(resolved):
                break

    return [
        [
            resolved.get(value[1]) if isinstance(value, tuple) and len(value) == 2
            and value[0] is _SHARED_STRING_REF else value
            for value in row
        ]
        for row in rows
    ]


def _xlsx_sheets(archive: zipfile.ZipFile) -> list[tuple[str, str]]:
    """Return (sheet name, worksheet XML path) pairs from an XLSX archive."""
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {
        relationship.attrib.get("Id", ""): relationship.attrib.get("Target", "")
        for relationship in relationships
        if _xml_name(relationship.tag) == "Relationship"
    }
    sheets: list[tuple[str, str]] = []
    for sheet in workbook.iter():
        if _xml_name(sheet.tag) != "sheet":
            continue
        relationship_id = next(
            (value for key, value in sheet.attrib.items() if _xml_name(key) == "id"), ""
        )
        target = targets.get(relationship_id, "").lstrip("/")
        if target and not target.startswith("xl/"):
            target = posixpath.normpath(posixpath.join("xl", target))
        if target:
            sheets.append((sheet.attrib.get("name", ""), target))
    return sheets


@lru_cache(maxsize=512)
def _inspect_xlsx_cached(
    path: str, size: int, mtime_ns: int,
) -> tuple[str, int, float, tuple[Any, ...]]:
    """Inspect an XLSX header without materialising its whole shared-string table."""
    del size, mtime_ns
    with zipfile.ZipFile(path) as archive:
        candidates: list[tuple[float, int, str, int, tuple[Any, ...]]] = []
        for sheet_name, sheet_path in _xlsx_sheets(archive):
            rows = _resolve_shared_strings(archive, _read_xlsx_rows(archive, sheet_path))
            header_row, confidence = choose_header_row(rows)
            header = tuple(rows[header_row]) if header_row < len(rows) else ()
            nonempty = sum(bool(_clean(value)) for value in header)
            candidates.append((confidence, nonempty, sheet_name, header_row, header))
        if not candidates:
            return "", 0, 0.0, ()
        candidates.sort(key=lambda item: (-item[0], -item[1]))
        confidence, _nonempty, sheet_name, header_row, header = candidates[0]
        return sheet_name, header_row, confidence, header


def xlsx_layout_and_header(path: str) -> tuple[str, int, float, list[Any]]:
    """Return the likely data sheet, zero-based header row, confidence, and values."""
    stat = os.stat(path)
    sheet, header_row, confidence, header = _inspect_xlsx_cached(
        os.path.abspath(path), stat.st_size, stat.st_mtime_ns,
    )
    return sheet, header_row, confidence, list(header)


def detect_xlsx_layout(path: str) -> tuple[str, int, float]:
    sheet, header_row, confidence, _header = xlsx_layout_and_header(path)
    return sheet, header_row, confidence


def detect_xls_layout(path: str) -> tuple[str | int, int, float]:
    import pandas as pd

    workbook = pd.ExcelFile(path)
    candidates: list[tuple[float, int, str, int]] = []
    for sheet_name in workbook.sheet_names:
        preview = pd.read_excel(
            workbook, sheet_name=sheet_name, header=None, nrows=HEADER_SCAN_ROWS + 1,
        )
        rows = preview.where(preview.notna(), None).values.tolist()
        header_row, confidence = choose_header_row(rows)
        nonempty = sum(bool(_clean(value)) for value in rows[header_row]) if rows else 0
        candidates.append((confidence, nonempty, sheet_name, header_row))
    if not candidates:
        return 0, 0, 0.0
    candidates.sort(key=lambda item: (-item[0], -item[1]))
    confidence, _, sheet_name, header_row = candidates[0]
    return sheet_name, header_row, confidence


def describe_layout(path: str) -> tuple[str, int, float]:
    extension = Path(path).suffix.lower()
    if extension == ".xlsx":
        return detect_xlsx_layout(path)
    if extension == ".xls":
        sheet, row, confidence = detect_xls_layout(path)
        return str(sheet), row, confidence
    return "", 0, 1.0
