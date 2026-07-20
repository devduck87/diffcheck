"""Canvas ベースの Excel 風グリッド表示。

行列入れ替え（転置）、差分色、同一値ハイライト、選択、行対応編集の
行選択表示に対応する。左右で1つずつ生成される。
"""

from __future__ import annotations

import tkinter as tk

from .model import col_letter

CW = 96      # セル幅
CH = 22      # セル高
HDR_W = 52   # 行ヘッダ幅
HDR_H = 24   # 列ヘッダ高
SINGLE_CW = 380   # 1列表示のセル幅（長文を折り返して表示）
SINGLE_PAD = 6    # 1列表示のセル内余白
FONT = ("Meiryo", 9)
HFONT = ("Meiryo", 9, "bold")

CHANGED_BG = "#ffe8b3"
REMOVED_BG = "#ffd6d6"
ADDED_BG = "#d6f5df"
SAME_BG = "#ffffff"
NONE_BG = "#ededed"
HEADER_BG = "#eef1f5"
GRID_LINE = "#cfcfcf"
SEL_OUTLINE = "#111111"
SAME_OUTLINE = "#1a73e8"
ROWSEL_BG = "#cfe3ff"


class GridView(tk.Frame):
    def __init__(self, parent, app, side: str):
        super().__init__(parent)
        self.app = app
        self.side = side  # "left" / "right"

        self.canvas = tk.Canvas(self, background="white", highlightthickness=0)
        self.vbar = tk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.hbar = tk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(
            yscrollcommand=self._on_yscroll,
            xscrollcommand=self._on_xscroll,
        )
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.vbar.grid(row=0, column=1, sticky="ns")
        self.hbar.grid(row=1, column=0, sticky="ew")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        self.canvas.bind("<Button-1>", self._on_click)
        self.canvas.bind("<B1-Motion>", self._on_drag_resize)
        self.canvas.bind("<ButtonRelease-1>", self._on_end_resize)
        self.canvas.bind("<Double-Button-1>", self._on_dbl)
        self.canvas.bind("<Motion>", self._on_motion)
        self.canvas.bind("<Button-3>", self._on_rclick)
        self.canvas.bind("<MouseWheel>", self._on_wheel)
        self.canvas.bind("<Shift-MouseWheel>", self._on_shift_wheel)

        # 行高・列幅ドラッグ用の状態と、直近レイアウトのオフセット
        self._resize = None          # (axis, index, start_coord, start_size)
        self._H = []                 # 横軸要素 [{kind,key,size}, ...]
        self._V = []                 # 縦軸要素
        self._hx = [HDR_W]           # 横境界の累積座標（len(H)+1）
        self._vy = [HDR_H]           # 縦境界の累積座標（len(V)+1）

    # ----------------------------------------------------------- スクロール
    def _on_yscroll(self, lo, hi):
        self.vbar.set(lo, hi)
        if self.app.should_sync("v"):
            self.app.propagate_scroll("v", lo, self)

    def _on_xscroll(self, lo, hi):
        self.hbar.set(lo, hi)
        if self.app.should_sync("h"):
            self.app.propagate_scroll("h", lo, self)

    def _on_wheel(self, event):
        self.canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")

    def _on_shift_wheel(self, event):
        self.canvas.xview_scroll(-1 if event.delta > 0 else 1, "units")

    # ----------------------------------------------------------- 座標変換
    def _sheet(self):
        return self.app.left_sheet if self.side == "left" else self.app.right_sheet

    def _pair_data_row(self, pair):
        return pair.left if self.side == "left" else pair.right

    def _build_axes(self):
        """現在の向き・可変サイズに基づき、横軸(H)・縦軸(V)と累積座標を作る。

        H/V の各要素は {kind:'field'|'slot', key:int, size:int}。
        行高(slot)は左右共通、列幅(field)は side ごと。
        """
        app = self.app
        pairs = app.model.pairs
        ncols = self._sheet().ncols
        cols = [{"kind": "field", "key": c, "size": app.field_extent(self.side, c)}
                for c in range(ncols)]
        slots = [{"kind": "slot", "key": k, "size": app.slot_extent(k)}
                 for k in range(len(pairs))]
        if app.transposed:
            self._H, self._V = slots, cols
        else:
            self._H, self._V = cols, slots
        self._hx = _accumulate(HDR_W, self._H)
        self._vy = _accumulate(HDR_H, self._V)

    def _hi_vi(self, k: int, col: int):
        """(スロット k, 列 col) を (横index, 縦index) へ。"""
        return (k, col) if self.app.transposed else (col, k)

    def _cellbox(self, k: int, col: int):
        hi, vi = self._hi_vi(k, col)
        return (self._hx[hi], self._vy[vi],
                self._H[hi]["size"], self._V[vi]["size"])

    def _hit(self, event):
        """クリック位置を (k, col) へ変換。範囲外は None。"""
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        if cx < HDR_W or cy < HDR_H:
            return None
        hi = _index_at(self._hx, cx)
        vi = _index_at(self._vy, cy)
        if hi is None or vi is None:
            return None
        k, col = (hi, vi) if self.app.transposed else (vi, hi)
        pairs = self.app.model.pairs
        sheet = self._sheet()
        if not (0 <= k < len(pairs)) or not (0 <= col < sheet.ncols):
            return None
        return (k, col)

    def _hit_single(self, event):
        """1列表示でのクリック位置を表示行 i へ変換。範囲外は None。"""
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        if cx < self.app.single_row_header_width() or cy < HDR_H:
            return None
        offs = self.app.single_offsets
        hs = self.app.single_heights
        for i, oy in enumerate(offs):
            if oy <= cy < oy + hs[i]:
                return i
        return None

    def _single_hit_data(self, event):
        i = self._hit_single(event)
        if i is None:
            return None
        info = self.app.single_cell(self.side, i)
        if info is None:
            return None
        pair, dr, col, state, text = info
        if dr is None:
            return None
        return (dr, col)

    # ----------------------------------------------------------- クリック
    def _on_click(self, event):
        if self._begin_resize(event):
            return
        if self.app.single_col:
            data = self._single_hit_data(event)
            if data is None:
                return
            dr, col = data
            if self.app.mode == "rowedit":
                self.app.on_row_select(self.side, dr)
            else:
                self.app.on_cell_select(self.side, dr, col)
            return
        hit = self._hit(event)
        if hit is None:
            return
        k, col = hit
        pair = self.app.model.pairs[k]
        dr = self._pair_data_row(pair)
        if dr is None:
            return
        if self.app.mode == "rowedit":
            self.app.on_row_select(self.side, dr)
        else:
            self.app.on_cell_select(self.side, dr, col)

    def _on_rclick(self, event):
        if self.app.single_col:
            data = self._single_hit_data(event)
            if data is None:
                return
            dr, _col = data
            self.app.show_context_menu(self.side, dr, event.x_root, event.y_root)
            return
        hit = self._hit(event)
        if hit is None:
            return
        k, col = hit
        pair = self.app.model.pairs[k]
        dr = self._pair_data_row(pair)
        if dr is None:
            return
        self.app.show_context_menu(self.side, dr, event.x_root, event.y_root)

    # ----------------------------------------------------- 行高・列幅リサイズ
    def _divider_at(self, cx, cy):
        """ヘッダ上の境界線近傍か判定。('H'|'V', 要素index) または None。"""
        if self.app.single_col:
            return None
        tol = 4
        if cy < HDR_H and cx >= HDR_W:      # 上ヘッダ → 横軸(H)の境界
            for i in range(1, len(self._hx)):
                if abs(cx - self._hx[i]) <= tol:
                    return ("H", i - 1)
        if cx < HDR_W and cy >= HDR_H:      # 左ヘッダ → 縦軸(V)の境界
            for i in range(1, len(self._vy)):
                if abs(cy - self._vy[i]) <= tol:
                    return ("V", i - 1)
        return None

    def _begin_resize(self, event) -> bool:
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        div = self._divider_at(cx, cy)
        if div is None:
            return False
        axis, idx = div
        elems = self._H if axis == "H" else self._V
        start_coord = cx if axis == "H" else cy
        self._resize = (axis, idx, start_coord, elems[idx]["size"])
        return True

    def _on_drag_resize(self, event):
        if self._resize is None:
            return
        axis, idx, start_coord, start_size = self._resize
        cur = self.canvas.canvasx(event.x) if axis == "H" else self.canvas.canvasy(event.y)
        new_size = int(start_size + (cur - start_coord))
        elems = self._H if axis == "H" else self._V
        el = elems[idx]
        self.app.set_extent(el["kind"], self.side, el["key"], new_size)

    def _on_end_resize(self, _event):
        self._resize = None

    def _on_motion(self, event):
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        div = self._divider_at(cx, cy)
        if div is None:
            self.canvas.configure(cursor="")
        elif div[0] == "H":
            self.canvas.configure(cursor="sb_h_double_arrow")
        else:
            self.canvas.configure(cursor="sb_v_double_arrow")

    def _on_dbl(self, event):
        """境界のダブルクリックで内容に合わせて自動調整。"""
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        div = self._divider_at(cx, cy)
        if div is None:
            return
        axis, idx = div
        el = (self._H if axis == "H" else self._V)[idx]
        self.app.autofit(el["kind"], self.side, el["key"])

    # ----------------------------------------------------------- 描画
    def redraw(self):
        if self.app.single_col:
            self._redraw_single()
            return
        c = self.canvas
        c.delete("all")
        model = self.app.model
        pairs = model.pairs
        sheet = self._sheet()
        n = len(pairs)
        ncols = sheet.ncols

        self._build_axes()
        width, height = self._hx[-1], self._vy[-1]
        c.configure(scrollregion=(0, 0, max(width, 1), max(height, 1)))

        sel = self.app.selection
        same = self.app.same_left if self.side == "left" else self.app.same_right
        rowsel = (self.app.left_sel_row if self.side == "left"
                  else self.app.right_sel_row)

        # セル本体
        for k, pair in enumerate(pairs):
            dr = self._pair_data_row(pair)
            row_selected = (self.app.mode == "rowedit" and dr is not None
                            and dr == rowsel)
            for col in range(ncols):
                x, y, w, h = self._cellbox(k, col)
                state = model.cell_state(self.side, pair, col)
                bg = _bg_for(state)
                if row_selected and state != "none":
                    bg = ROWSEL_BG
                c.create_rectangle(x, y, x + w, y + h, fill=bg,
                                   outline=GRID_LINE, width=1)
                if dr is not None and state != "none":
                    text = sheet.text(dr, col)
                    if text:
                        # 上揃え・1行に切り詰めてセル境界からはみ出させない
                        disp = self._fit_text(text, w - 8)
                        c.create_text(x + 4, y + 3, anchor="nw", text=disp,
                                      font=FONT)
                # 同一値ハイライト（枠線）
                if dr is not None and (dr, col) in same:
                    c.create_rectangle(x + 1, y + 1, x + w - 1, y + h - 1,
                                       outline=SAME_OUTLINE, width=2)
                # 選択セル（太枠）
                if (sel is not None and sel[0] == self.side
                        and sel[1] == dr and sel[2] == col and dr is not None):
                    c.create_rectangle(x + 1, y + 1, x + w - 1, y + h - 1,
                                       outline=SEL_OUTLINE, width=3)

        self._draw_headers(n, ncols, pairs, sheet)

    def _elem_label(self, el, horizontal: bool) -> str:
        """H/V 要素の見出し文字列。"""
        if el["kind"] == "field":
            return col_letter(el["key"])
        pair = self.app.model.pairs[el["key"]]
        dr = self._pair_data_row(pair)
        label = str(dr + 1) if dr is not None else "-"
        mark = self._kind_mark(pair)
        return f"{mark}{label}行" if horizontal else f"{mark}{label}"

    def _draw_headers(self, n, ncols, pairs, sheet):
        c = self.canvas
        # 上ヘッダ（横軸 H）
        for i, el in enumerate(self._H):
            x, w = self._hx[i], el["size"]
            c.create_rectangle(x, 0, x + w, HDR_H, fill=HEADER_BG,
                               outline=GRID_LINE)
            c.create_text(x + w / 2, HDR_H / 2, text=self._elem_label(el, True),
                          font=HFONT, width=w - 2)
        # 左ヘッダ（縦軸 V）
        for i, el in enumerate(self._V):
            y, h = self._vy[i], el["size"]
            c.create_rectangle(0, y, HDR_W, y + h, fill=HEADER_BG,
                               outline=GRID_LINE)
            c.create_text(HDR_W / 2, y + h / 2, text=self._elem_label(el, False),
                          font=HFONT, width=HDR_W - 2)
        # 左上コーナー
        c.create_rectangle(0, 0, HDR_W, HDR_H, fill="#dfe4ea", outline=GRID_LINE)
        c.create_text(HDR_W / 2, HDR_H / 2,
                      text=("左" if self.side == "left" else "右"), font=HFONT)

    def _kind_mark(self, pair) -> str:
        """行対応種別の簡易マーク。"""
        if pair.manual and pair.left is not None and pair.right is not None:
            return "🔗"
        return ""

    def _fit_text(self, text: str, max_w: int) -> str:
        """セル幅 max_w に収まる1行へ切り詰める（超過分は末尾を…に）。"""
        text = text.replace("\n", " ")
        f = self.app._font
        if max_w <= 0:
            return ""
        if f.measure(text) <= max_w:
            return text
        ell = "…"
        ew = f.measure(ell)
        out = []
        w = 0
        for chn in text:
            cwn = f.measure(chn)
            if w + cwn + ew > max_w:
                break
            out.append(chn)
            w += cwn
        return "".join(out) + ell

    # ----------------------------------------------------------- 1列表示 描画
    def _redraw_single(self):
        c = self.canvas
        c.delete("all")
        app = self.app
        hdr_w = app.single_row_header_width()
        cw = SINGLE_CW
        nd = len(app.single_offsets)
        total = app.single_total
        c.configure(scrollregion=(0, 0, hdr_w + cw, max(total, 1)))

        sel = app.selection
        same = app.same_left if self.side == "left" else app.same_right
        rowsel = (app.left_sel_row if self.side == "left"
                  else app.right_sel_row)

        for i in range(nd):
            y = app.single_offsets[i]
            h = app.single_heights[i]
            info = app.single_cell(self.side, i)
            if info is None:
                continue
            pair, dr, col, state, text = info
            bg = _bg_for(state)
            if app.mode == "rowedit" and dr is not None and dr == rowsel \
                    and state != "none":
                bg = ROWSEL_BG
            c.create_rectangle(hdr_w, y, hdr_w + cw, y + h, fill=bg,
                               outline=GRID_LINE)
            if dr is not None and text:
                c.create_text(hdr_w + SINGLE_PAD, y + 3, anchor="nw", text=text,
                              font=FONT, width=cw - 2 * SINGLE_PAD)
            if dr is not None and (dr, col) in same:
                c.create_rectangle(hdr_w + 1, y + 1, hdr_w + cw - 1, y + h - 1,
                                   outline=SAME_OUTLINE, width=2)
            if (sel is not None and sel[0] == self.side and dr is not None
                    and sel[1] == dr and sel[2] == col):
                c.create_rectangle(hdr_w + 1, y + 1, hdr_w + cw - 1, y + h - 1,
                                   outline=SEL_OUTLINE, width=3)
            # 行ヘッダ
            mark = self._kind_mark(pair) if not app.transposed else ""
            c.create_rectangle(0, y, hdr_w, y + h, fill=HEADER_BG,
                               outline=GRID_LINE)
            c.create_text(hdr_w / 2, y + h / 2,
                          text=f"{mark}{app.single_row_header_label(i)}",
                          font=HFONT, width=hdr_w - 4)

        # 列ヘッダ（表示中の列の見出し）
        c.create_rectangle(hdr_w, 0, hdr_w + cw, HDR_H, fill=HEADER_BG,
                           outline=GRID_LINE)
        c.create_text(hdr_w + cw / 2, HDR_H / 2,
                      text=app.single_col_header_label(), font=HFONT)
        # 左上コーナー
        c.create_rectangle(0, 0, hdr_w, HDR_H, fill="#dfe4ea", outline=GRID_LINE)
        c.create_text(hdr_w / 2, HDR_H / 2,
                      text=("左" if self.side == "left" else "右"), font=HFONT)


def _bg_for(state: str) -> str:
    return {
        "changed": CHANGED_BG,
        "removed": REMOVED_BG,
        "added": ADDED_BG,
        "same": SAME_BG,
        "none": NONE_BG,
    }.get(state, SAME_BG)


def _accumulate(start: int, elems: list) -> list:
    """要素サイズの累積境界座標を返す（長さ len(elems)+1）。"""
    offs = [start]
    acc = start
    for el in elems:
        acc += el["size"]
        offs.append(acc)
    return offs


def _index_at(offsets: list, coord: float):
    """累積境界 offsets 上で coord が入る区間 index を返す。範囲外は None。"""
    if coord < offsets[0] or coord >= offsets[-1]:
        return None
    lo, hi = 0, len(offsets) - 1
    while lo < hi - 1:
        mid = (lo + hi) // 2
        if offsets[mid] <= coord:
            lo = mid
        else:
            hi = mid
    return lo
