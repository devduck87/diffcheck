"""データモデル: Cell / Sheet / Workbook と共通ユーティリティ。"""

from __future__ import annotations


# 比較パネル（最大3ファイル）。先頭 "left" が基準ファイル。
# 内部キーは 2 ファイル時代の名前を引き継ぎ、画面表示は A / B / C を使う。
SIDES = ("left", "right", "third")
SIDE_LABELS = {"left": "A", "right": "B", "third": "C"}


def side_index(side) -> int:
    """パネルキー（"left" 等）または index を index へ正規化する。"""
    if isinstance(side, str):
        return SIDES.index(side)
    return int(side)


def side_label(side) -> str:
    return SIDE_LABELS[SIDES[side_index(side)]]


def col_letter(index: int) -> str:
    """0始まりの列番号を Excel の列文字（A, B, ..., Z, AA ...）へ変換する。"""
    if index < 0:
        return ""
    letters = ""
    n = index + 1
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(ord("A") + rem) + letters
    return letters


def cell_address(row: int, col: int) -> str:
    """0始まりの (row, col) を Excel セルアドレス（例: C5）へ変換する。"""
    return f"{col_letter(col)}{row + 1}"


class Cell:
    """1つのセル。

    value   : 内部値（数値なら int/float、文字列なら str、空なら ""）
    display : 画面表示用文字列
    formula : 数式文字列（無ければ None）
    """

    __slots__ = ("value", "display", "formula")

    def __init__(self, value="", display=None, formula=None):
        self.value = value
        self.display = display if display is not None else _to_display(value)
        self.formula = formula

    def text(self) -> str:
        """一致判定・表示に使う代表文字列。"""
        return self.display if self.display is not None else ""

    def is_blank(self) -> bool:
        return self.display is None or self.display == ""

    def __repr__(self):
        return f"Cell({self.display!r})"


def _to_display(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float):
        # 整数値の float は "1000.0" ではなく "1000" とする
        if value.is_integer():
            return str(int(value))
        return repr(value)
    return str(value)


EMPTY_CELL = Cell("")


class Sheet:
    """1シート。rows は Cell の2次元リスト（矩形にパディング済み）。"""

    def __init__(self, name: str, rows: list[list[Cell]]):
        self.name = name
        self.rows = rows
        self.nrows = len(rows)
        self.ncols = max((len(r) for r in rows), default=0)
        # 矩形化
        for r in self.rows:
            while len(r) < self.ncols:
                r.append(Cell(""))

    def cell(self, r: int, c: int) -> Cell:
        if 0 <= r < self.nrows and 0 <= c < self.ncols:
            return self.rows[r][c]
        return EMPTY_CELL

    def text(self, r: int, c: int) -> str:
        return self.cell(r, c).text()

    def __repr__(self):
        return f"Sheet({self.name!r}, {self.nrows}x{self.ncols})"


class Workbook:
    def __init__(self, path: str, sheets: list[Sheet]):
        self.path = path
        self.sheets = sheets

    def sheet_names(self) -> list[str]:
        return [s.name for s in self.sheets]

    def get(self, name: str) -> Sheet | None:
        for s in self.sheets:
            if s.name == name:
                return s
        return None
