"""同一値セルの索引（仕様 3.12）。

比較実行時に正規化済み値 -> セル位置一覧 のインデックスを作成し、
セル選択時に高速に同一値セルを検索する。2〜3ファイルに対応する。
"""

from __future__ import annotations

from .model import Sheet
from .normalize import NormalizeOptions, norm_key


class ValueIndex:
    def __init__(self, sheets, opts: NormalizeOptions):
        if isinstance(sheets, Sheet):
            raise TypeError(
                "ValueIndex はシートのリストを受け取ります"
                "（例: ValueIndex([left, right], opts)）")
        self.opts = opts
        self._maps: list[dict[str, list[tuple[int, int]]]] = []
        for sheet in sheets:
            self._maps.append(self._build(sheet))

    def _build(self, sheet: Sheet) -> dict:
        target: dict[str, list[tuple[int, int]]] = {}
        for r in range(sheet.nrows):
            for c in range(sheet.ncols):
                key = norm_key(sheet.text(r, c), self.opts)
                if key == "":  # 空白・空文字列は索引しない（仕様 3.7）
                    continue
                target.setdefault(key, []).append((r, c))
        return target

    def find(self, key: str) -> list[list[tuple[int, int]]]:
        """正規化キーに一致するセル位置を、パネルごとのリストで返す。

        戻り値: [[(row, col), ...],  # 1番目のシート
                 [(row, col), ...],  # 2番目のシート
                 ...]
        """
        if not key:
            return [[] for _ in self._maps]
        return [list(m.get(key, [])) for m in self._maps]
