"""差分エンジン: 行整列（difflib）+ セル比較 + 手動行対応 + Undo/Redo。

仕様 2.6 / 4 章に対応。
2ファイル比較に加えて 3ファイル比較にも対応する。3ファイルの場合は
先頭シート（パネルA）を基準とし、「A↔B」「A↔C」の整列結果を A の行位置で
マージして1本の整列スロット列にする。
"""

from __future__ import annotations

import difflib

from .model import SIDES, Sheet, side_index
from .normalize import NormalizeOptions, norm_key


# 行対応状態（仕様 4.3）
AUTO = "auto"        # 自動対応
MANUAL = "manual"    # 手動対応
LEFT_ONLY = "left_only"    # 基準のみ（削除行）
RIGHT_ONLY = "right_only"  # 比較先のみ（追加行）
PARTIAL = "partial"        # 3ファイル比較で一部のパネルにしか無い行


class RowSlot:
    """整列後の1行スロット。

    rows[i] は i 番目のシート（パネル）の行番号(0始まり)、無ければ None。
    2ファイル時代の left / right 属性も別名として使える。
    """

    __slots__ = ("rows", "manual", "changed_cols", "status")

    def __init__(self, rows, manual=False):
        self.rows: list = list(rows)
        self.manual = manual
        self.changed_cols: set[int] = set()
        self.status = ""  # equal / changed / left_only / right_only / partial

    # ------------------------------------------------------------- 別名
    @property
    def left(self):
        return self.rows[0]

    @left.setter
    def left(self, v):
        self.rows[0] = v

    @property
    def right(self):
        return self.rows[1] if len(self.rows) > 1 else None

    @right.setter
    def right(self, v):
        self.rows[1] = v

    @property
    def third(self):
        return self.rows[2] if len(self.rows) > 2 else None

    def row(self, side) -> int | None:
        i = side_index(side)
        return self.rows[i] if i < len(self.rows) else None

    def present(self) -> list[int]:
        """行を持つパネルの index 一覧。"""
        return [i for i, r in enumerate(self.rows) if r is not None]

    def snapshot(self):
        return (tuple(self.rows), self.manual)

    @classmethod
    def restore(cls, snap):
        return cls(snap[0], snap[1])

    def kind(self) -> str:
        """対応種別（表示・アイコン用）。"""
        p = self.present()
        if len(p) == len(self.rows):
            return MANUAL if self.manual else AUTO
        if len(self.rows) == 2:
            return LEFT_ONLY if 0 in p else RIGHT_ONLY
        return PARTIAL

    def __repr__(self):
        return f"RowSlot({self.rows}, {self.status})"


# 旧名（2ファイル時代の呼び名）
RowPair = RowSlot


class DiffModel:
    """2〜3シートの差分状態を保持する。

        DiffModel([left, right], opts)          # 2ファイル
        DiffModel([left, right, third], opts)   # 3ファイル（基準=left）
    """

    def __init__(self, sheets, opts: NormalizeOptions):
        if isinstance(sheets, Sheet):
            raise TypeError(
                "DiffModel はシートのリストを受け取ります"
                "（例: DiffModel([left, right], opts)）")
        self.sheets: list[Sheet] = list(sheets)
        if not 2 <= len(self.sheets) <= len(SIDES):
            raise ValueError(f"比較できるのは2〜{len(SIDES)}ファイルです。")
        self.opts = opts
        self.pairs: list[RowSlot] = []
        self._undo: list[list] = []
        self._redo: list[list] = []
        self._keys: list[list[tuple]] = []
        self._build_keys()
        self.rebuild_auto()

    # ------------------------------------------------------------------ 別名
    @property
    def npanes(self) -> int:
        return len(self.sheets)

    @property
    def left(self) -> Sheet:
        return self.sheets[0]

    @property
    def right(self) -> Sheet:
        return self.sheets[1]

    def sheet(self, side) -> Sheet:
        return self.sheets[side_index(side)]

    # ---------------------------------------------------------------- キー生成
    def _build_keys(self):
        self._keys = [[self._row_key(sh, r) for r in range(sh.nrows)]
                      for sh in self.sheets]

    def _row_key(self, sheet: Sheet, r: int) -> tuple:
        return tuple(norm_key(sheet.text(r, c), self.opts)
                     for c in range(sheet.ncols))

    def cell_key(self, side, r: int, c: int) -> str:
        return norm_key(self.sheet(side).text(r, c), self.opts)

    # ---------------------------------------------------------------- 自動整列
    def rebuild_auto(self):
        """基準シートと各シートを difflib で整列し直す（自動対応）。"""
        n = self.npanes
        nbase = self.sheets[0].nrows
        matches: list[dict] = []   # パネル j: 基準行 -> j の行
        orphans: list[dict] = []   # パネル j: 挿入位置(基準行index) -> [j の行, ...]
        for j in range(1, n):
            m, o = _align_two(self._keys[0], self._keys[j])
            matches.append(m)
            orphans.append(o)

        slots: list[RowSlot] = []
        for a in range(nbase + 1):   # nbase は末尾（基準行より後ろ）の位置
            pend = [orphans[j][a] if a in orphans[j] else []
                    for j in range(n - 1)]
            slots.extend(self._merge_orphans(pend))
            if a < nbase:
                rows = [a] + [matches[j].get(a) for j in range(n - 1)]
                slots.append(RowSlot(rows))
        self.pairs = slots
        self._recompute_all()

    def _merge_orphans(self, pend: list[list[int]]) -> list[RowSlot]:
        """基準に無い行（各パネルの余り行）をスロットにする。

        3ファイル比較では、B と C の両方に同じ内容で増えた行を1スロットに
        まとめる（同じ追加行が2段に分かれて見えるのを防ぐ）。
        """
        n = self.npanes
        if n == 2:
            return [RowSlot([None, b]) for b in pend[0]]
        slots = []
        used = [set() for _ in pend]
        for j, rows_j in enumerate(pend):
            for b in rows_j:
                if b in used[j]:
                    continue
                used[j].add(b)
                rows = [None] * n
                rows[j + 1] = b
                key = self._keys[j + 1][b]
                for j2 in range(j + 1, len(pend)):
                    for b2 in pend[j2]:
                        if b2 in used[j2]:
                            continue
                        if self._keys[j2 + 1][b2] == key:
                            rows[j2 + 1] = b2
                            used[j2].add(b2)
                            break
                slots.append(RowSlot(rows))
        return slots

    # ---------------------------------------------------------------- セル比較
    def _recompute_all(self):
        for p in self.pairs:
            self._recompute(p)

    def _recompute(self, p: RowSlot):
        present = p.present()
        keys = [self._keys[i][p.rows[i]] for i in present]
        changed = set()
        if len(keys) >= 2:
            ncols = max(len(k) for k in keys)
            for c in range(ncols):
                vals = {k[c] if c < len(k) else "" for k in keys}
                if len(vals) > 1:
                    changed.add(c)
        p.changed_cols = changed
        if len(present) == self.npanes:
            p.status = "equal" if not changed else "changed"
        elif self.npanes == 2:
            p.status = "left_only" if 0 in present else "right_only"
        else:
            p.status = PARTIAL

    def cell_state(self, side, slot: RowSlot, col: int) -> str:
        """描画用のセル状態: same / changed / added / removed / none。"""
        i = side_index(side)
        if i >= len(slot.rows) or slot.rows[i] is None:
            return "none"
        if col in slot.changed_cols:
            return "changed"
        if slot.status in ("equal", "changed"):
            return "same"
        # 一部のパネルにしか無い行：基準にあれば「削除」、無ければ「追加」
        return "removed" if slot.rows[0] is not None else "added"

    # ---------------------------------------------------------------- 手動対応
    def slot_of(self, side, row: int):
        """パネル side の行 row を含むスロット index（無ければ None）。"""
        i = side_index(side)
        for k, p in enumerate(self.pairs):
            if i < len(p.rows) and p.rows[i] == row:
                return k
        return None

    def is_paired(self, side_a, row_a: int, side_b, row_b: int) -> bool:
        ka = self.slot_of(side_a, row_a)
        return ka is not None and ka == self.slot_of(side_b, row_b)

    def partner_of(self, side_from, row: int, side_to):
        """行 row と同じスロットにある side_to の行を返す（無ければ None）。"""
        k = self.slot_of(side_from, row)
        if k is None:
            return None
        return self.pairs[k].row(side_to)

    def manual_pair(self, side_a, row_a: int, side_b, row_b: int):
        """side_b の行 row_b を、side_a の行 row_a のスロットへ移す。

        1パネル1行の対応を保つため、玉突き（スワップ）方式で入れ替える。
        """
        ia, ib = side_index(side_a), side_index(side_b)
        if ia == ib:
            return
        ka = self.slot_of(side_a, row_a)
        kb = self.slot_of(side_b, row_b)
        if ka is None or kb is None or ka == kb:
            return
        self._push_undo()
        pa, pb = self.pairs[ka], self.pairs[kb]
        pa.rows[ib], pb.rows[ib] = row_b, pa.rows[ib]
        pa.manual = True
        pb.manual = len(pb.present()) >= 2
        if not pb.present():
            self.pairs.pop(kb)
        else:
            self._recompute(pb)
        self._recompute(pa)
        self._redo.clear()

    def unpair(self, side, row: int):
        """行 row を含むスロットの対応を解除し、パネルごとに分割する（仕様 4.8）。"""
        k = self.slot_of(side, row)
        if k is None:
            return
        p = self.pairs[k]
        present = p.present()
        if len(present) <= 1:
            return  # すでに単独
        self._push_undo()
        n = self.npanes
        new_slots = []
        for i in present[1:]:
            rows = [None] * n
            rows[i] = p.rows[i]
            p.rows[i] = None
            ns = RowSlot(rows)
            self._recompute(ns)
            new_slots.append(ns)
        p.manual = False
        self._recompute(p)
        for off, ns in enumerate(new_slots):
            self.pairs.insert(k + 1 + off, ns)
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
        self.pairs = [RowSlot.restore(s) for s in snap]
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
        """追加＝基準に無い行、削除＝基準にあり他に無い行、変更＝値違い。"""
        added = removed = changed = equal = 0
        changed_cells = 0
        for p in self.pairs:
            present = p.present()
            changed_cells += len(p.changed_cols)
            if p.rows[0] is None:
                added += 1
            elif len(present) < self.npanes:
                removed += 1
            elif p.changed_cols:
                changed += 1
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


def _align_two(akeys: list, bkeys: list) -> tuple[dict, dict]:
    """基準 akeys と bkeys を整列し、(対応, 余り行) を返す。

    対応  : {基準行 -> b の行}
    余り行: {挿入位置(基準行index。末尾は len(akeys)) -> [b の行, ...]}
    """
    sm = difflib.SequenceMatcher(None, akeys, bkeys, autojunk=False)
    match: dict[int, int] = {}
    orphans: dict[int, list[int]] = {}
    pending: list[int] = []

    def flush(pos: int):
        if pending:
            orphans.setdefault(pos, []).extend(pending)
            pending.clear()

    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            for k in range(i2 - i1):
                flush(i1 + k)
                match[i1 + k] = j1 + k
        elif tag == "replace":
            n = min(i2 - i1, j2 - j1)
            for k in range(n):
                flush(i1 + k)
                match[i1 + k] = j1 + k
            # 基準側の余り（i1+n 〜 i2）は b に対応が無い行なので記録しない
            for k in range(n, j2 - j1):
                pending.append(j1 + k)
        elif tag == "insert":
            for k in range(j1, j2):
                pending.append(k)
        # delete（基準のみの行）は match に載せないだけでよい
    flush(len(akeys))
    return match, orphans
