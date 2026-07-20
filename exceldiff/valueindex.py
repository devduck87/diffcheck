"""同一値セルの索引（仕様 3.12）。

比較実行時に正規化済み値 -> セル位置一覧 のインデックスを作成し、
セル選択時に高速に同一値セルを検索する。
"""

from __future__ import annotations

from .model import Sheet
from .normalize import NormalizeOptions, norm_key


class ValueIndex:
    def __init__(self, left: Sheet, right: Sheet, opts: NormalizeOptions):
        self.opts = opts
        self._left_map: dict[str, list[tuple[int, int]]] = {}
        self._right_map: dict[str, list[tuple[int, int]]] = {}
        self._build(left, "left")
        self._build(right, "right")

    def _build(self, sheet: Sheet, side: str):
        target = self._left_map if side == "left" else self._right_map
        for r in range(sheet.nrows):
            for c in range(sheet.ncols):
                key = norm_key(sheet.text(r, c), self.opts)
                if key == "":  # 空白・空文字列は索引しない（仕様 3.7）
                    continue
                target.setdefault(key, []).append((r, c))

    def find(self, key: str):
        """正規化キーに一致する左右のセル位置を返す。

        戻り値: (left_cells, right_cells) 各々 [(row, col), ...]
        """
        if not key:
            return [], []
        return (
            list(self._left_map.get(key, [])),
            list(self._right_map.get(key, [])),
        )
