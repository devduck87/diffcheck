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
        self.canvas.bind("<Button-3>", self._on_rclick)
        self.canvas.bind("<MouseWheel>", self._on_wheel)
        self.canvas.bind("<Shift-MouseWheel>", self._on_shift_wheel)

    # ----------------------------------------------------------- スクロール
    def _on_yscroll(self, lo, hi):
        self.vbar.set(lo, hi)
        axis = "v"
        if self.app.aligned_axis() == axis:
            self.app.propagate_scroll(axis, lo, self)

    def _on_xscroll(self, lo, hi):
        self.hbar.set(lo, hi)
        axis = "h"
        if self.app.aligned_axis() == axis:
            self.app.propagate_scroll(axis, lo, self)

    def _on_wheel(self, event):
        self.canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")

    def _on_shift_wheel(self, event):
        self.canvas.xview_scroll(-1 if event.delta > 0 else 1, "units")

    # ----------------------------------------------------------- 座標変換
    def _sheet(self):
        return self.app.left_sheet if self.side == "left" else self.app.right_sheet

    def _pair_data_row(self, pair):
        return pair.left if self.side == "left" else pair.right

    def _cellbox(self, k: int, col: int):
        """(整列スロット k, シート列 col) の描画矩形 (x, y, w, h)。"""
        if self.app.transposed:
            return (HDR_W + k * CW, HDR_H + col * CH, CW, CH)
        return (HDR_W + col * CW, HDR_H + k * CH, CW, CH)

    def _hit(self, event):
        """クリック位置を (k, col) へ変換。範囲外は None。"""
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        if cx < HDR_W or cy < HDR_H:
            return None
        if self.app.transposed:
            k = int((cx - HDR_W) // CW)
            col = int((cy - HDR_H) // CH)
        else:
            col = int((cx - HDR_W) // CW)
            k = int((cy - HDR_H) // CH)
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

        if self.app.transposed:
            width = HDR_W + n * CW
            height = HDR_H + ncols * CH
        else:
            width = HDR_W + ncols * CW
            height = HDR_H + n * CH
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
                        c.create_text(x + 4, y + h / 2, anchor="w", text=text,
                                      font=FONT, width=w - 6)
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

    def _draw_headers(self, n, ncols, pairs, sheet):
        c = self.canvas
        if self.app.transposed:
            # 上ヘッダ = 整列スロット（元の行番号）、左ヘッダ = 列文字
            for k, pair in enumerate(pairs):
                dr = self._pair_data_row(pair)
                x = HDR_W + k * CW
                label = str(dr + 1) if dr is not None else "-"
                mark = self._kind_mark(pair)
                c.create_rectangle(x, 0, x + CW, HDR_H, fill=HEADER_BG,
                                   outline=GRID_LINE)
                c.create_text(x + CW / 2, HDR_H / 2, text=f"{mark}{label}行",
                              font=HFONT)
            for col in range(ncols):
                y = HDR_H + col * CH
                c.create_rectangle(0, y, HDR_W, y + CH, fill=HEADER_BG,
                                   outline=GRID_LINE)
                c.create_text(HDR_W / 2, y + CH / 2, text=col_letter(col),
                              font=HFONT)
        else:
            # 上ヘッダ = 列文字、左ヘッダ = 元の行番号
            for col in range(ncols):
                x = HDR_W + col * CW
                c.create_rectangle(x, 0, x + CW, HDR_H, fill=HEADER_BG,
                                   outline=GRID_LINE)
                c.create_text(x + CW / 2, HDR_H / 2, text=col_letter(col),
                              font=HFONT)
            for k, pair in enumerate(pairs):
                dr = self._pair_data_row(pair)
                y = HDR_H + k * CH
                label = str(dr + 1) if dr is not None else "-"
                mark = self._kind_mark(pair)
                c.create_rectangle(0, y, HDR_W, y + CH, fill=HEADER_BG,
                                   outline=GRID_LINE)
                c.create_text(HDR_W / 2, y + CH / 2, text=f"{mark}{label}",
                              font=HFONT)
        # 左上コーナー
        c.create_rectangle(0, 0, HDR_W, HDR_H, fill="#dfe4ea", outline=GRID_LINE)
        c.create_text(HDR_W / 2, HDR_H / 2,
                      text=("左" if self.side == "left" else "右"), font=HFONT)

    def _kind_mark(self, pair) -> str:
        """行対応種別の簡易マーク。"""
        if pair.manual and pair.left is not None and pair.right is not None:
            return "🔗"
        return ""

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
