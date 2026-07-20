"""差分エンジン: 行整列（difflib）+ セル比較 + 手動行対応 + Undo/Redo。

仕様 2.6 / 4 章に対応。
"""

from __future__ import annotations

import difflib

from .model import Sheet
from .normalize import NormalizeOptions, norm_key


# 行対応状態（仕様 4.3）
AUTO = "auto"        # 自動対応
MANUAL = "manual"    # 手動対応
LEFT_ONLY = "left_only"    # 左のみ（削除行）
RIGHT_ONLY = "right_only"  # 右のみ（追加行）


class RowPair:
    """整列後の1行スロット。left/right は各シートの行番号(0始まり)または None。"""

    __slots__ = ("left", "right", "manual", "changed_cols", "status")

    def __init__(self, left, right, manual=False):
        self.left = left
        self.right = right
        self.manual = manual
        self.changed_cols: set[int] = set()
        self.status = ""  # equal / changed / left_only / right_only

    def snapshot(self):
        return (self.left, self.right, self.manual)

    @classmethod
    def restore(cls, snap):
        return cls(snap[0], snap[1], snap[2])

    def kind(self) -> str:
        """対応種別（表示・アイコン用）。"""
        if self.left is not None and self.right is not None:
            return MANUAL if self.manual else AUTO
        if self.left is not None:
            return LEFT_ONLY
        return RIGHT_ONLY


class DiffModel:
    """左右2シートの差分状態を保持する。"""

    def __init__(self, left: Sheet, right: Sheet, opts: NormalizeOptions):
        self.left = left
        self.right = right
        self.opts = opts
        self.pairs: list[RowPair] = []
        self._undo: list[list] = []
        self._redo: list[list] = []
        self._lkeys: list[tuple] = []
        self._rkeys: list[tuple] = []
        self._build_keys()
        self.rebuild_auto()

    # ---------------------------------------------------------------- キー生成
    def _build_keys(self):
        self._lkeys = [self._row_key(self.left, r) for r in range(self.left.nrows)]
        self._rkeys = [self._row_key(self.right, r) for r in range(self.right.nrows)]

    def _row_key(self, sheet: Sheet, r: int) -> tuple:
        return tuple(norm_key(sheet.text(r, c), self.opts) for c in range(sheet.ncols))

    def cell_key(self, side: str, r: int, c: int) -> str:
        sheet = self.left if side == "left" else self.right
        return norm_key(sheet.text(r, c), self.opts)

    # ---------------------------------------------------------------- 自動整列
    def rebuild_auto(self):
        """difflib で左右の行を整列し直す（自動対応）。"""
        sm = difflib.SequenceMatcher(None, self._lkeys, self._rkeys, autojunk=False)
        pairs: list[RowPair] = []
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag == "equal":
                for k in range(i2 - i1):
                    pairs.append(RowPair(i1 + k, j1 + k))
            elif tag == "replace":
                n = min(i2 - i1, j2 - j1)
                for k in range(n):
                    pairs.append(RowPair(i1 + k, j1 + k))
                for k in range(n, i2 - i1):
                    pairs.append(RowPair(i1 + k, None))
                for k in range(n, j2 - j1):
                    pairs.append(RowPair(None, j1 + k))
            elif tag == "delete":
                for k in range(i1, i2):
                    pairs.append(RowPair(k, None))
            elif tag == "insert":
                for k in range(j1, j2):
                    pairs.append(RowPair(None, k))
        self.pairs = pairs
        self._recompute_all()

    # ---------------------------------------------------------------- セル比較
    def _recompute_all(self):
        for p in self.pairs:
            self._recompute(p)

    def _recompute(self, p: RowPair):
        if p.left is not None and p.right is not None:
            lk = self._lkeys[p.left]
            rk = self._rkeys[p.right]
            ncols = max(len(lk), len(rk))
            changed = set()
            for c in range(ncols):
                a = lk[c] if c < len(lk) else ""
                b = rk[c] if c < len(rk) else ""
                if a != b:
                    changed.add(c)
            p.changed_cols = changed
            p.status = "equal" if not changed else "changed"
        elif p.left is not None:
            p.status = "left_only"
            p.changed_cols = set()
        else:
            p.status = "right_only"
            p.changed_cols = set()

    def cell_state(self, side: str, pair: RowPair, col: int) -> str:
        """描画用のセル状態: same / changed / added / removed / none。"""
        if side == "left":
            if pair.left is None:
                return "none"
            if pair.right is None:
                return "removed"
            return "changed" if col in pair.changed_cols else "same"
        else:
            if pair.right is None:
                return "none"
            if pair.left is None:
                return "added"
            return "changed" if col in pair.changed_cols else "same"

    # ---------------------------------------------------------------- 手動対応
    def _find_left(self, L):
        for i, p in enumerate(self.pairs):
            if p.left == L:
                return i
        return None

    def _find_right(self, R):
        for i, p in enumerate(self.pairs):
            if p.right == R:
                return i
        return None

    def is_paired(self, L, R) -> bool:
        i = self._find_left(L)
        return i is not None and self.pairs[i].right == R

    def existing_partner_of_right(self, R):
        """右行 R が現在対応している左行を返す（無ければ None）。"""
        i = self._find_right(R)
        if i is None:
            return None
        return self.pairs[i].left

    def manual_pair(self, L: int, R: int):
        """左行 L と右行 R を手動対応させる（スワップ方式で1:1を維持）。"""
        iL = self._find_left(L)
        iR = self._find_right(R)
        if iL is None or iR is None:
            return
        if iL == iR:  # すでに対応済み
            return
        self._push_undo()
        pL = self.pairs[iL]
        pR = self.pairs[iR]
        oldR = pL.right   # L が元々対応していた右行
        oldL = pR.left    # R が元々対応していた左行
        # L <-> R を確定
        pL.right = R
        pL.manual = True
        # 玉突きで oldL <-> oldR を組む
        pR.left = oldL
        pR.right = oldR
        pR.manual = bool(oldL is not None and oldR is not None)
        if pR.left is None and pR.right is None:
            self.pairs.pop(iR)
        else:
            self._recompute(pR)
        self._recompute(pL)
        self._redo.clear()

    def unpair(self, L=None, R=None):
        """指定行の対応を解除し、左のみ/右のみに分割する（仕様 4.8）。"""
        idx = None
        if L is not None:
            idx = self._find_left(L)
        elif R is not None:
            idx = self._find_right(R)
        if idx is None:
            return
        p = self.pairs[idx]
        if p.left is None or p.right is None:
            return  # すでに片側のみ
        self._push_undo()
        r = p.right
        p.right = None
        p.manual = False
        self._recompute(p)
        newp = RowPair(None, r)
        self._recompute(newp)
        self.pairs.insert(idx + 1, newp)
        self._redo.clear()

    def restore_auto_all(self):
        """全体を自動対応へ戻す（仕様 4.9）。"""
        self._push_undo()
        self.rebuild_auto()
        self._redo.clear()

    # ---------------------------------------------------------------- Undo/Redo
    def _clone(self) -> list:
        return [p.snapshot() for p in self.pairs]

    def _push_undo(self):
        self._undo.append(self._clone())
        if len(self._undo) > 200:
            self._undo.pop(0)

    def _apply(self, snap: list):
        self.pairs = [RowPair.restore(s) for s in snap]
        self._recompute_all()

    def can_undo(self) -> bool:
        return bool(self._undo)

    def can_redo(self) -> bool:
        return bool(self._redo)

    def undo(self):
        if not self._undo:
            return
        self._redo.append(self._clone())
        self._apply(self._undo.pop())

    def redo(self):
        if not self._redo:
            return
        self._undo.append(self._clone())
        self._apply(self._redo.pop())

    # ---------------------------------------------------------------- 集計
    def summary(self) -> dict:
        added = removed = changed = equal = 0
        changed_cells = 0
        for p in self.pairs:
            if p.status == "left_only":
                removed += 1
            elif p.status == "right_only":
                added += 1
            elif p.status == "changed":
                changed += 1
                changed_cells += len(p.changed_cols)
            else:
                equal += 1
        return {
            "added": added,
            "removed": removed,
            "changed": changed,
            "equal": equal,
            "changed_cells": changed_cells,
            "manual": sum(1 for p in self.pairs if p.manual),
        }
