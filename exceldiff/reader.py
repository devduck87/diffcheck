"""Excel(.xlsx / .xlsm) / CSV 読込。標準ライブラリのみ。

.xlsx / .xlsm は zip + XML なので zipfile と xml.etree で解析する（openpyxl 不使用）。
.xlsm はマクロを含むだけで中身の構造は .xlsx と同じなので同じパーサで読む
（マクロ本体 vbaProject.bin は比較対象外）。
数値書式（表示形式）は解釈せず、内部値をそのまま表示する簡易実装。
"""

from __future__ import annotations

import csv
import os
import zipfile
import xml.etree.ElementTree as ET

from .model import Cell, Sheet, Workbook


def _local(tag: str) -> str:
    """名前空間付きタグからローカル名だけを取り出す。"""
    return tag.rsplit("}", 1)[-1]


def load(path: str) -> Workbook:
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        return _load_xlsx(path)
    if ext in (".csv", ".txt", ".tsv"):
        return _load_csv(path, "\t" if ext == ".tsv" else ",")
    raise ValueError(f"未対応の拡張子です: {ext}")


# --------------------------------------------------------------------------- CSV
def _load_csv(path: str, delimiter: str) -> Workbook:
    rows: list[list[Cell]] = []
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        for record in csv.reader(f, delimiter=delimiter):
            rows.append([Cell(v) for v in record])
    name = os.path.splitext(os.path.basename(path))[0]
    return Workbook(path, [Sheet(name, rows)])


# --------------------------------------------------------------------------- XLSX
def _load_xlsx(path: str) -> Workbook:
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        shared = _read_shared_strings(z) if "xl/sharedStrings.xml" in names else []
        sheet_map = _read_workbook_sheets(z)   # [(name, target_path)]
        sheets = []
        for name, target in sheet_map:
            full = "xl/" + target if not target.startswith("xl/") else target
            if full not in names:
                # 相対パスの調整（worksheets/sheet1.xml など）
                cand = "xl/" + target.lstrip("/")
                full = cand if cand in names else full
            if full not in names:
                continue
            rows = _read_worksheet(z.read(full), shared)
            sheets.append(Sheet(name, rows))
    if not sheets:
        raise ValueError("シートが見つかりませんでした。")
    return Workbook(path, sheets)


def _read_shared_strings(z: zipfile.ZipFile) -> list[str]:
    root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    result = []
    for si in root:
        if _local(si.tag) != "si":
            continue
        result.append(_extract_text(si))
    return result


def _extract_text(si_elem) -> str:
    """<si> 要素からテキストを連結して取り出す（run <r><t> にも対応）。"""
    parts = []
    for node in si_elem.iter():
        if _local(node.tag) == "t" and node.text:
            parts.append(node.text)
    return "".join(parts)


def _read_workbook_sheets(z: zipfile.ZipFile) -> list[tuple[str, str]]:
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = _read_rels(z)
    result = []
    for elem in wb.iter():
        if _local(elem.tag) != "sheet":
            continue
        name = elem.get("name", "Sheet")
        rid = None
        for k, v in elem.attrib.items():
            if _local(k) == "id":  # r:id
                rid = v
                break
        target = rels.get(rid, "worksheets/sheet1.xml")
        result.append((name, target))
    return result


def _read_rels(z: zipfile.ZipFile) -> dict[str, str]:
    path = "xl/_rels/workbook.xml.rels"
    if path not in z.namelist():
        return {}
    root = ET.fromstring(z.read(path))
    out = {}
    for rel in root:
        rid = rel.get("Id")
        target = rel.get("Target", "")
        if rid:
            out[rid] = target
    return out


def _read_worksheet(data: bytes, shared: list[str]) -> list[list[Cell]]:
    root = ET.fromstring(data)
    rows_by_index: dict[int, dict[int, Cell]] = {}
    max_col = 0

    for elem in root.iter():
        if _local(elem.tag) != "row":
            continue
        r_attr = elem.get("r")
        row_idx = int(r_attr) - 1 if r_attr else len(rows_by_index)
        cells: dict[int, Cell] = {}
        for c in elem:
            if _local(c.tag) != "c":
                continue
            ref = c.get("r")
            col_idx = _col_from_ref(ref) if ref else 0
            cell = _parse_cell(c, shared)
            cells[col_idx] = cell
            if col_idx > max_col:
                max_col = col_idx
        rows_by_index[row_idx] = cells

    if not rows_by_index:
        return []

    max_row = max(rows_by_index)
    result = []
    for r in range(max_row + 1):
        cells = rows_by_index.get(r, {})
        row = [cells.get(c, Cell("")) for c in range(max_col + 1)]
        result.append(row)
    return result


def _parse_cell(c_elem, shared: list[str]) -> Cell:
    ctype = c_elem.get("t")  # s, str, b, e, inlineStr, None(number)
    formula = None
    value_text = None
    for child in c_elem:
        tag = _local(child.tag)
        if tag == "f":
            formula = child.text
        elif tag == "v":
            value_text = child.text
        elif tag == "is":  # inline string
            value_text = _extract_text(child)
            ctype = "inlineStr"

    if value_text is None and ctype != "inlineStr":
        return Cell("", formula=("=" + formula if formula else None))

    if ctype == "s":  # shared string index
        try:
            text = shared[int(value_text)]
        except (ValueError, IndexError, TypeError):
            text = value_text or ""
        return Cell(text, display=text, formula=("=" + formula if formula else None))

    if ctype in ("str", "inlineStr", "e"):
        text = value_text or ""
        return Cell(text, display=text, formula=("=" + formula if formula else None))

    if ctype == "b":
        text = "TRUE" if value_text == "1" else "FALSE"
        return Cell(text, display=text, formula=("=" + formula if formula else None))

    # 数値
    try:
        num = float(value_text)
        value = int(num) if num.is_integer() else num
    except (ValueError, TypeError):
        value = value_text or ""
    return Cell(value, formula=("=" + formula if formula else None))


def _col_from_ref(ref: str) -> int:
    letters = ""
    for ch in ref:
        if ch.isalpha():
            letters += ch
        else:
            break
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch.upper()) - ord("A") + 1)
    return idx - 1
