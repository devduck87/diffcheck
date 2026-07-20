"""メインアプリケーション（tkinter）。

3つの追加機能を統合する:
  * 行列入れ替え表示（左右同時）
  * 同一値セルハイライト
  * 手動での行対応修正（Undo/Redo 対応）
"""

from __future__ import annotations

import os
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk

from . import reader
from .diffengine import DiffModel
from .gridview import GridView, CW, CH, HDR_H, GRID_LINE, SINGLE_CW, SINGLE_PAD
from .model import Sheet, cell_address, col_letter
from .normalize import NormalizeOptions
from .valueindex import ValueIndex

MID_W = 66
MAX_HIGHLIGHT = 100  # 同一値ハイライトの上限（仕様 3.9）

# 1列表示モードの行ヘッダ幅
SINGLE_HDR_W_NORMAL = 52   # 通常向き＝元の行番号
SINGLE_HDR_W_TRANS = 120   # 転置＝フィールド名を表示


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel差分比較ツール")
        self.geometry("1280x760")

        # --- 状態 -----------------------------------------------------------
        self.left_sheet: Sheet | None = None
        self.right_sheet: Sheet | None = None
        self.left_wb = None
        self.right_wb = None
        self.model: DiffModel | None = None
        self.vindex: ValueIndex | None = None

        self.transposed = False
        self.highlight_enabled = tk.BooleanVar(value=False)
        self.show_lines = tk.BooleanVar(value=True)
        self.mode = "normal"  # normal / rowedit

        # 1列表示モード（1列だけを幅広で表示して左右比較する）
        self.single_col = False
        self.single_index = 0            # 表示中の列（現在の向きの列軸インデックス）
        self.single_offsets: list = []   # 各表示行のy開始位置（可変高）
        self.single_heights: list = []
        self.single_total = 0
        self._font = tkfont.Font(self, family="Meiryo", size=9)
        self._line_h = self._font.metrics("linespace")
        self._wrap_px = SINGLE_CW - 2 * SINGLE_PAD

        self.selection = None  # (side, r, c)
        self.same_left: set = set()
        self.same_right: set = set()
        self.left_sel_row = None
        self.right_sel_row = None
        self._same_nav: list = []   # 同一値移動用
        self._same_nav_i = -1

        # 正規化条件（差分比較と同一値ハイライトで共用: 仕様 3.5）
        self.opts = NormalizeOptions(
            trim=True, fullwidth=False, number_equiv=True, blank_equiv=True
        )
        self._num_opts = {
            "ignore_case": tk.BooleanVar(value=False),
            "trim": tk.BooleanVar(value=True),
            "fullwidth": tk.BooleanVar(value=False),
            "number_equiv": tk.BooleanVar(value=True),
        }

        self._syncing = False

        # セル寸法（可変）。#4 の自動フィットは未実装のため全セル一律。
        self.cell_w = CW
        self.cell_h = CH
        # 左右のスクロール位置を（整列軸に加えて）両軸で一致させるか
        self.sync_both = tk.BooleanVar(value=False)

        self._build_menu()
        self._build_toolbar()
        self._build_detail_panel()
        self._build_body()
        self._build_rowedit_panel()
        self._build_statusbar()

        self._refresh_buttons()

    # ================================================================= UI構築
    def _build_menu(self):
        m = tk.Menu(self)
        self.config(menu=m)

        fm = tk.Menu(m, tearoff=0)
        fm.add_command(label="左ファイルを開く...", command=lambda: self.open_file("left"))
        fm.add_command(label="右ファイルを開く...", command=lambda: self.open_file("right"))
        fm.add_separator()
        fm.add_command(label="終了", command=self.destroy)
        m.add_cascade(label="ファイル", menu=fm)

        cm = tk.Menu(m, tearoff=0)
        cm.add_command(label="比較を再実行", command=self.recompare)
        m.add_cascade(label="比較", menu=cm)

        vm = tk.Menu(m, tearoff=0)
        vm.add_command(label="行と列を入れ替える / 通常表示に戻す",
                       command=self.toggle_transpose)
        vm.add_command(label="1列表示 / 通常のグリッド表示",
                       command=self.toggle_single)
        vm.add_checkbutton(label="同一値セルをハイライトする",
                           variable=self.highlight_enabled,
                           command=self.on_highlight_toggle)
        vm.add_checkbutton(label="行対応関係を線で表示する",
                           variable=self.show_lines, command=self.redraw_all)
        vm.add_checkbutton(label="左右のスクロール位置を一致させる",
                           variable=self.sync_both, command=self._on_sync_toggle)
        m.add_cascade(label="表示", menu=vm)

        rm = tk.Menu(m, tearoff=0)
        rm.add_command(label="行対応編集モードの切替", command=self.toggle_mode)
        rm.add_command(label="選択した行を対応付ける", command=self.do_pair)
        rm.add_command(label="行対応を解除する", command=self.do_unpair)
        rm.add_command(label="自動対応に戻す（全体）", command=self.do_restore_auto)
        m.add_cascade(label="行対応", menu=rm)

        sm = tk.Menu(m, tearoff=0)
        for key, label in [
            ("ignore_case", "大文字・小文字を無視する"),
            ("trim", "前後の空白を無視する"),
            ("fullwidth", "全角・半角を同一視する"),
            ("number_equiv", "数値と数値文字列を同一視する"),
        ]:
            sm.add_checkbutton(label=label, variable=self._num_opts[key],
                               command=self.recompare)
        m.add_cascade(label="設定", menu=sm)

    def _build_toolbar(self):
        bar = tk.Frame(self, bd=1, relief="raised")
        bar.pack(side="top", fill="x")

        def btn(text, cmd):
            b = tk.Button(bar, text=text, command=cmd, padx=6)
            b.pack(side="left", padx=2, pady=3)
            return b

        self.b_prev = btn("◀ 前差分", lambda: self.goto_diff(-1))
        self.b_next = btn("次差分 ▶", lambda: self.goto_diff(1))
        tk.Frame(bar, width=2, bg="#bbb").pack(side="left", fill="y", padx=4)
        self.b_transpose = btn("⇄ 行列入替", self.toggle_transpose)
        self.b_single = btn("1列表示", self.toggle_single)
        self.b_colprev = btn("◀列", lambda: self.step_column(-1))
        self.b_colnext = btn("列▶", lambda: self.step_column(1))
        self.b_highlight = btn("同一値強調", self.toggle_highlight_btn)
        tk.Frame(bar, width=2, bg="#bbb").pack(side="left", fill="y", padx=4)
        tk.Label(bar, text="幅").pack(side="left")
        self.sp_w = tk.Spinbox(bar, from_=40, to=400, width=4, increment=4,
                               command=self._on_cell_size)
        self.sp_w.delete(0, "end")
        self.sp_w.insert(0, str(self.cell_w))
        self.sp_w.pack(side="left", padx=(0, 4))
        tk.Label(bar, text="高").pack(side="left")
        self.sp_h = tk.Spinbox(bar, from_=14, to=200, width=4, increment=2,
                               command=self._on_cell_size)
        self.sp_h.delete(0, "end")
        self.sp_h.insert(0, str(self.cell_h))
        self.sp_h.pack(side="left", padx=(0, 2))
        for sp in (self.sp_w, self.sp_h):
            sp.bind("<Return>", self._on_cell_size)
            sp.bind("<FocusOut>", self._on_cell_size)
        tk.Frame(bar, width=2, bg="#bbb").pack(side="left", fill="y", padx=4)
        self.b_mode = btn("行対応編集", self.toggle_mode)
        self.b_pair = btn("対応付け", self.do_pair)
        self.b_unpair = btn("対応解除", self.do_unpair)
        self.b_restore = btn("自動復元", self.do_restore_auto)
        tk.Frame(bar, width=2, bg="#bbb").pack(side="left", fill="y", padx=4)
        self.b_undo = btn("Undo", self.do_undo)
        self.b_redo = btn("Redo", self.do_redo)

    def _build_detail_panel(self):
        """差分一覧・左グリッド・右グリッドの上に、選択セルの値詳細を表示する帯。

        値が長い場合に備え、左右の値はスクロール可能な Text で表示する。
        """
        p = tk.Frame(self, bd=1, relief="sunken", height=110)
        p.pack(side="top", fill="x")
        p.pack_propagate(False)

        # 差分一覧の上：見出しと選択セル情報（スクロール可能）
        info = tk.Frame(p, width=230)
        info.pack(side="left", fill="y")
        info.pack_propagate(False)
        tk.Label(info, text="セル詳細", bg="#dfe4ea", anchor="w").pack(fill="x")
        self.detail_addr = self._make_scroll_text(info)

        # 左グリッドの上：左の値（全文・スクロール可能）
        lf = tk.Frame(p)
        lf.pack(side="left", fill="both", expand=True)
        tk.Label(lf, text="左の値", bg="#eef1f5", anchor="w").pack(fill="x")
        self.detail_left = self._make_scroll_text(lf)

        # 中央（対応列）の上：スペーサ
        tk.Frame(p, width=MID_W).pack(side="left", fill="y")

        # 右グリッドの上：右の値（全文・スクロール可能）
        rf = tk.Frame(p)
        rf.pack(side="left", fill="both", expand=True)
        tk.Label(rf, text="右の値", bg="#eef1f5", anchor="w").pack(fill="x")
        self.detail_right = self._make_scroll_text(rf)

        self._clear_detail()

    def _make_scroll_text(self, parent) -> tk.Text:
        """縦スクロールバー付きの読み取り専用 Text を作って返す。"""
        wrap = tk.Frame(parent)
        wrap.pack(fill="both", expand=True)
        sb = tk.Scrollbar(wrap, orient="vertical")
        sb.pack(side="right", fill="y")
        txt = tk.Text(wrap, wrap="word", font=("Meiryo", 9), height=1,
                      bd=0, padx=4, pady=2, background="white",
                      yscrollcommand=sb.set, state="disabled",
                      cursor="arrow")
        txt.pack(side="left", fill="both", expand=True)
        sb.config(command=txt.yview)
        # ホイールでスクロール（フォーカス不要でホバー中に効かせる）
        txt.bind("<MouseWheel>",
                 lambda e, t=txt: (t.yview_scroll(-1 if e.delta > 0 else 1,
                                                  "units"), "break")[1])
        return txt

    @staticmethod
    def _set_text(widget: tk.Text, text: str):
        widget.config(state="normal")
        widget.delete("1.0", tk.END)
        if text:
            widget.insert("1.0", text)
        widget.config(state="disabled")
        widget.yview_moveto(0)

    def _update_detail(self, side, r, c):
        """選択セルと同じ整列行の左右の値を詳細帯に表示する。"""
        if not self.model:
            return
        lr = rr = None
        for p in self.model.pairs:
            dr = p.left if side == "left" else p.right
            if dr == r:
                lr, rr = p.left, p.right
                break
        lval = (self.left_sheet.text(lr, c)
                if lr is not None and c < self.left_sheet.ncols else "")
        rval = (self.right_sheet.text(rr, c)
                if rr is not None and c < self.right_sheet.ncols else "")
        self._set_text(self.detail_addr,
                       f"列：{self._col_label(c)}\n"
                       f"選択：{side} {cell_address(r, c)}")
        self._set_text(self.detail_left, lval)
        self._set_text(self.detail_right, rval)

    def _clear_detail(self):
        if hasattr(self, "detail_addr"):
            self._set_text(self.detail_addr, "セル未選択")
            self._set_text(self.detail_left, "")
            self._set_text(self.detail_right, "")

    def _on_cell_size(self, *_):
        try:
            w = int(self.sp_w.get())
            h = int(self.sp_h.get())
        except (ValueError, TypeError):
            return
        w = max(40, min(400, w))
        h = max(14, min(200, h))
        if (w, h) == (self.cell_w, self.cell_h):
            return
        self.cell_w, self.cell_h = w, h
        self.redraw_all()

    def _on_sync_toggle(self):
        # 有効化した瞬間に、左の現在位置へ右（と中央）を合わせる
        if self.sync_both.get():
            self.propagate_scroll("h", self.left_grid.canvas.xview()[0],
                                  self.left_grid)
            self.propagate_scroll("v", self.left_grid.canvas.yview()[0],
                                  self.left_grid)

    def should_sync(self, axis: str) -> bool:
        if self.sync_both.get():
            return True
        return self.aligned_axis() == axis

    def _build_body(self):
        body = tk.Frame(self)
        body.pack(side="top", fill="both", expand=True)

        # 差分一覧
        left_panel = tk.Frame(body, width=230)
        left_panel.pack(side="left", fill="y")
        left_panel.pack_propagate(False)
        tk.Label(left_panel, text="差分一覧", bg="#dfe4ea").pack(fill="x")
        self.difflist = tk.Listbox(left_panel, font=("Meiryo", 9),
                                   activestyle="none")
        self.difflist.pack(fill="both", expand=True)
        self.difflist.bind("<<ListboxSelect>>", self._on_difflist_select)

        # 左グリッド
        self.left_grid = GridView(body, self, "left")
        self.left_grid.pack(side="left", fill="both", expand=True)

        # 中央（対応関係）
        self.mid = tk.Canvas(body, width=MID_W, background="#f7f8fa",
                             highlightthickness=0)
        self.mid.pack(side="left", fill="y")

        # 右グリッド
        self.right_grid = GridView(body, self, "right")
        self.right_grid.pack(side="left", fill="both", expand=True)

    def _build_rowedit_panel(self):
        p = tk.Frame(self, bd=1, relief="sunken")
        p.pack(side="top", fill="x")
        self.rowedit_label = tk.Label(
            p, anchor="w",
            text="通常モード：セルを選択すると詳細・同一値を表示します。",
            font=("Meiryo", 9))
        self.rowedit_label.pack(side="left", padx=8, pady=3)

    def _build_statusbar(self):
        sb = tk.Frame(self, bd=1, relief="sunken")
        sb.pack(side="bottom", fill="x")
        self.status_cell = tk.Label(sb, text="元セル：-", width=22, anchor="w")
        self.status_cell.pack(side="left", padx=4)
        self.status_same = tk.Label(sb, text="同一値：-", width=24, anchor="w")
        self.status_same.pack(side="left", padx=4)
        self.status_map = tk.Label(sb, text="対応：-", width=24, anchor="w")
        self.status_map.pack(side="left", padx=4)
        self.status_sum = tk.Label(sb, text="集計：-", anchor="w")
        self.status_sum.pack(side="left", padx=4)

    # ================================================================= ファイル
    def open_file(self, side: str):
        path = filedialog.askopenfilename(
            title=f"{side} ファイルを開く",
            filetypes=[("Excel/CSV", "*.xlsx *.csv *.tsv *.txt"), ("すべて", "*.*")])
        if not path:
            return
        try:
            wb = reader.load(path)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("読込エラー", f"{os.path.basename(path)}\n{e}")
            return
        sheet = wb.sheets[0]
        if side == "left":
            self.left_wb, self.left_sheet = wb, sheet
        else:
            self.right_wb, self.right_sheet = wb, sheet
        if self.left_sheet and self.right_sheet:
            self.recompare()

    def load_pair(self, left_path: str, right_path: str):
        self.left_wb = reader.load(left_path)
        self.right_wb = reader.load(right_path)
        self.left_sheet = self.left_wb.sheets[0]
        self.right_sheet = self.right_wb.sheets[0]
        self.recompare()

    def recompare(self):
        if not (self.left_sheet and self.right_sheet):
            return
        for k, var in self._num_opts.items():
            setattr(self.opts, k, var.get())
        self.model = DiffModel(self.left_sheet, self.right_sheet, self.opts)
        self.vindex = ValueIndex(self.left_sheet, self.right_sheet, self.opts)
        self.selection = None
        self.same_left.clear()
        self.same_right.clear()
        self.left_sel_row = self.right_sel_row = None
        self._clear_detail()
        self._rebuild_difflist()
        self.redraw_all()
        self._update_summary()
        self._refresh_buttons()

    # ================================================================= 描画
    def redraw_all(self):
        if not self.model:
            return
        if self.single_col:
            self.build_single_layout()
        self.left_grid.redraw()
        self.right_grid.redraw()
        self._draw_mid()

    def _draw_mid(self):
        c = self.mid
        c.delete("all")
        if not self.model:
            return
        # 1列表示（通常向き）では可変高で対応線を描く。転置1列は1レコードなので省略。
        if self.single_col:
            if self.transposed:
                c.configure(scrollregion=(0, 0, MID_W, MID_W))
                c.create_text(MID_W / 2, 40, text="1列\n表示中", font=("Meiryo", 8),
                              fill="#888", justify="center")
                return
            c.configure(scrollregion=(0, 0, MID_W, max(self.single_total, 1)))
            if not self.show_lines.get():
                return
            for k, p in enumerate(self.model.pairs):
                if k >= len(self.single_offsets):
                    break
                y = self.single_offsets[k] + self.single_heights[k] / 2
                self._draw_mid_connector(c, p, y)
            return
        n = len(self.model.pairs)
        height = HDR_H + n * self.cell_h
        c.configure(scrollregion=(0, 0, MID_W, max(height, 1)))
        if self.transposed:
            c.create_text(MID_W / 2, 40, text="行列\n入替中", font=("Meiryo", 8),
                          fill="#888", justify="center")
            return
        if not self.show_lines.get():
            return
        for k, p in enumerate(self.model.pairs):
            y = HDR_H + k * self.cell_h + self.cell_h / 2
            self._draw_mid_connector(c, p, y)

    def _draw_mid_connector(self, c, p, y):
        if p.left is not None and p.right is not None:
            if p.manual:
                c.create_line(4, y, MID_W - 4, y, fill="#1a73e8", width=3)
                c.create_text(MID_W / 2, y - 7, text="🔗", font=("Meiryo", 7))
            else:
                c.create_line(6, y, MID_W - 6, y, fill="#9aa0a6", width=1)
        elif p.left is not None:
            c.create_text(MID_W / 2, y, text="◀削除", font=("Meiryo", 7),
                          fill="#c0392b")
        else:
            c.create_text(MID_W / 2, y, text="追加▶", font=("Meiryo", 7),
                          fill="#1e824c")

    # ---------------------------------------------------------- 1列表示レイアウト
    def _single_ndisp(self) -> int:
        """1列表示の表示行数（現在の向きの行軸）。"""
        if self.transposed:
            return max(self.left_sheet.ncols, self.right_sheet.ncols)
        return len(self.model.pairs)

    def single_cols_count(self) -> int:
        """1列表示で選択できる列の総数（現在の向きの列軸）。"""
        if self.transposed:
            return len(self.model.pairs)
        return max(self.left_sheet.ncols, self.right_sheet.ncols)

    def single_cell(self, side: str, i: int):
        """表示行 i の (pair, data_row, col, state, text) を返す。範囲外は None。"""
        pairs = self.model.pairs
        j0 = self.single_index
        if self.transposed:
            if not (0 <= j0 < len(pairs)):
                return None
            pair = pairs[j0]
            col = i
        else:
            if not (0 <= i < len(pairs)):
                return None
            pair = pairs[i]
            col = j0
        dr = pair.left if side == "left" else pair.right
        state = self.model.cell_state(side, pair, col)
        sheet = self.left_sheet if side == "left" else self.right_sheet
        text = sheet.text(dr, col) if (dr is not None and col < sheet.ncols) else ""
        return (pair, dr, col, state, text)

    def _wrap_lines(self, text: str) -> int:
        """折り返し後の行数を推定する。"""
        if not text:
            return 1
        total = 0
        for para in text.split("\n"):
            if para == "":
                total += 1
                continue
            line_w = 0
            lines = 1
            for ch in para:
                w = self._font.measure(ch)
                if line_w + w > self._wrap_px and line_w > 0:
                    lines += 1
                    line_w = w
                else:
                    line_w += w
            total += lines
        return total

    def build_single_layout(self):
        """左右で共通の可変行高を計算し、対応行を揃える。"""
        self.single_offsets = []
        self.single_heights = []
        self.single_total = 0
        if not self.model:
            return
        nd = self._single_ndisp()
        # 選択列を範囲内へ丸める
        cc = self.single_cols_count()
        if cc > 0:
            self.single_index = max(0, min(self.single_index, cc - 1))
        y = HDR_H
        min_h = self._line_h + SINGLE_PAD
        for i in range(nd):
            lt = self.single_cell("left", i)
            rt = self.single_cell("right", i)
            ll = self._wrap_lines(lt[4]) if lt else 1
            rl = self._wrap_lines(rt[4]) if rt else 1
            lines = max(ll, rl, 1)
            h = max(min_h, lines * self._line_h + SINGLE_PAD)
            self.single_offsets.append(y)
            self.single_heights.append(h)
            y += h
        self.single_total = y

    def single_row_header_width(self) -> int:
        return SINGLE_HDR_W_TRANS if self.transposed else SINGLE_HDR_W_NORMAL

    def single_col_header_label(self) -> str:
        """1列表示ヘッダ（表示中の列が何か）のラベル。"""
        j0 = self.single_index
        if self.transposed:
            # 列＝1つの整列スロット（1レコード）
            if 0 <= j0 < len(self.model.pairs):
                p = self.model.pairs[j0]
                lt = f"左{p.left + 1}" if p.left is not None else "左-"
                rt = f"右{p.right + 1}" if p.right is not None else "右-"
                return f"{lt} / {rt} 行"
            return "-"
        # 列＝1つのフィールド。見出し行があれば列名を使う。
        name = self.left_sheet.text(0, j0) if self.left_sheet.nrows > 0 else ""
        return name or col_letter(j0)

    def single_row_header_label(self, i: int) -> str:
        """1列表示の各表示行の見出し。"""
        if self.transposed:
            # 行＝フィールド。見出し行の値を使う。
            name = self.left_sheet.text(0, i) if self.left_sheet.nrows > 0 else ""
            return name or col_letter(i)
        # 行＝整列スロット（元の行番号）。左右で異なるので side ごとに出す方が正確だが
        # 中央寄せの簡易表示として左優先→なければ右。
        if 0 <= i < len(self.model.pairs):
            p = self.model.pairs[i]
            if p.left is not None:
                return str(p.left + 1)
            if p.right is not None:
                return str(p.right + 1)
        return "-"

    # ---------------------------------------------------------- 1列表示 操作
    def toggle_single(self):
        if not self.model:
            self.single_col = not self.single_col
            return
        self.single_col = not self.single_col
        self.b_single.config(relief="sunken" if self.single_col else "raised")
        if self.single_col:
            self.single_index = self._derive_single_index()
        self.redraw_all()
        self._refresh_buttons()
        self._update_single_status()

    def _derive_single_index(self) -> int:
        """現在の選択から表示する列を決める。"""
        if self.selection:
            side, r, c = self.selection
            if self.transposed:
                for k, p in enumerate(self.model.pairs):
                    dr = p.left if side == "left" else p.right
                    if dr == r:
                        return k
            else:
                return c
        return 0

    def step_column(self, direction: int):
        if not (self.model and self.single_col):
            return
        cc = self.single_cols_count()
        if cc <= 0:
            return
        self.single_index = (self.single_index + direction) % cc
        self.redraw_all()
        self._update_single_status()

    def _update_single_status(self):
        if self.single_col and self.model:
            cc = self.single_cols_count()
            axis = "レコード" if self.transposed else "フィールド"
            self.status_map.config(
                text=f"1列表示：{axis} {self.single_index + 1}/{cc}"
                     f"（{self.single_col_header_label()}）")
        else:
            self._update_summary()

    # ---------------------------------------------------------- スクロール同期
    def aligned_axis(self) -> str:
        # 1列表示は常に表示行を縦に積むため縦同期。
        if self.single_col:
            return "v"
        return "h" if self.transposed else "v"

    def propagate_scroll(self, axis: str, lo, source):
        if self._syncing:
            return
        self._syncing = True
        try:
            src_canvas = getattr(source, "canvas", source)
            targets = [self.left_grid.canvas, self.right_grid.canvas]
            if axis == "v":
                targets.append(self.mid)
            for t in targets:
                if t is src_canvas:
                    continue
                if axis == "v":
                    t.yview_moveto(lo)
                else:
                    t.xview_moveto(lo)
        finally:
            self._syncing = False

    def _ensure_visible(self, k: int, col: int):
        if not self.model:
            return
        if self.single_col:
            # 表示行 i を割り出して縦位置へスクロール
            i = k if not self.transposed else col
            if 0 <= i < len(self.single_offsets) and self.single_total > 0:
                frac = max(0, self.single_offsets[i] / self.single_total - 0.05)
                self.left_grid.canvas.yview_moveto(frac)
                self.propagate_scroll("v", frac, self.left_grid)
            return
        n = len(self.model.pairs)
        ncols = (self.left_sheet.ncols if self.left_sheet else 1)
        ch = self.cell_h
        cw = self.cell_w
        if self.transposed:
            total = HDR_H + max(ncols, 1) * ch
            self.left_grid.canvas.yview_moveto(max(0, (HDR_H + col * ch) / total - 0.1))
            self.right_grid.canvas.yview_moveto(max(0, (HDR_H + col * ch) / total - 0.1))
            frac = (k * cw) / max(1, n * cw)
            self.left_grid.canvas.xview_moveto(max(0, frac - 0.1))
            self.propagate_scroll("h", max(0, frac - 0.1), self.left_grid)
        else:
            total = HDR_H + max(n, 1) * ch
            frac = max(0, (HDR_H + k * ch) / total - 0.1)
            self.left_grid.canvas.yview_moveto(frac)
            self.propagate_scroll("v", frac, self.left_grid)

    # ================================================================= 選択
    def on_cell_select(self, side: str, r: int, c: int):
        self.selection = (side, r, c)
        self._update_same_value(side, r, c)
        self._update_cell_status(side, r, c)
        self._update_detail(side, r, c)
        self.redraw_all()

    def _update_same_value(self, side, r, c):
        self.same_left.clear()
        self.same_right.clear()
        self._same_nav = []
        self._same_nav_i = -1
        if not self.highlight_enabled.get() or not self.vindex:
            self.status_same.config(text="同一値：-")
            return
        sheet = self.left_sheet if side == "left" else self.right_sheet
        cell = sheet.cell(r, c)
        key = self.model.cell_key(side, r, c)
        if cell.is_blank() or key == "":
            self.status_same.config(text="同一値：空白は対象外")
            return
        lc, rc = self.vindex.find(key)
        total = len(lc) + len(rc)
        capped = ""
        if total > MAX_HIGHLIGHT:
            keep = MAX_HIGHLIGHT
            lc2 = lc[: max(0, min(len(lc), keep))]
            rc2 = rc[: max(0, keep - len(lc2))]
            lc, rc = lc2, rc2
            capped = f"（先頭{MAX_HIGHLIGHT}件を表示）"
        self.same_left = set(lc)
        self.same_right = set(rc)
        self._same_nav = ([("left", *p) for p in lc] + [("right", *p) for p in rc])
        self.status_same.config(text=f"同一値：左{len(lc)}件 右{len(rc)}件{capped}")

    def _update_cell_status(self, side, r, c):
        self.status_cell.config(text=f"元セル：{cell_address(r, c)}（{side}）")

    # ================================================================= 行対応編集
    def on_row_select(self, side: str, dr: int):
        if side == "left":
            self.left_sel_row = dr
        else:
            self.right_sel_row = dr
        self._update_rowedit_label()
        self.redraw_all()

    def _update_rowedit_label(self):
        lt = f"{self.left_sel_row + 1}行目" if self.left_sel_row is not None else "-"
        rt = f"{self.right_sel_row + 1}行目" if self.right_sel_row is not None else "-"
        self.rowedit_label.config(
            text=f"行対応編集モード｜左選択行：{lt}　右選択行：{rt}"
                 f"　→［対応付け］で左右を対応させます")

    def do_pair(self):
        if not self.model:
            return
        if self.mode != "rowedit":
            messagebox.showinfo("行対応編集", "「行対応編集」モードに切り替えてください。")
            return
        if self.transposed:
            messagebox.showwarning(
                "行列入れ替え中",
                "行列入れ替え表示中は行対応を編集できません。\n通常表示へ戻してください。")
            return
        L, R = self.left_sel_row, self.right_sel_row
        if L is None or R is None:
            messagebox.showwarning(
                "選択不足",
                "左行または右行が選択されていません。\n左右から1行ずつ選択してください。")
            return
        if self.model.is_paired(L, R):
            messagebox.showinfo("行対応", "選択した行はすでに対応付けられています。")
            return
        partner = self.model.existing_partner_of_right(R)
        if partner is not None and partner != L:
            ok = messagebox.askyesno(
                "対応の競合",
                f"右{R + 1}行目は、すでに左{partner + 1}行目と対応付けられています。\n\n"
                f"既存の対応を解除して、左{L + 1}行目と対応付けますか？")
            if not ok:
                return
        self.model.manual_pair(L, R)
        self._after_map_change()

    def do_unpair(self):
        if not self.model:
            return
        if self.mode != "rowedit":
            messagebox.showinfo("行対応編集", "「行対応編集」モードに切り替えてください。")
            return
        L = self.left_sel_row
        R = self.right_sel_row
        if L is None and R is None:
            messagebox.showwarning("選択不足", "解除する行を選択してください。")
            return
        self.model.unpair(L=L, R=R)
        self._after_map_change()

    def do_restore_auto(self):
        if not self.model:
            return
        if messagebox.askyesno("自動対応に戻す",
                               "手動対応をすべて破棄し、自動対応へ戻しますか？"):
            self.model.restore_auto_all()
            self._after_map_change()

    def _after_map_change(self):
        self._rebuild_difflist()
        self._reeval_same_value()
        self.redraw_all()
        self._update_summary()
        self._refresh_buttons()

    # ================================================================= モード
    def toggle_mode(self):
        if self.mode == "normal":
            if self.transposed:
                messagebox.showwarning(
                    "行列入れ替え中",
                    "行列入れ替え表示中は行対応を編集できません。\n通常表示へ戻してください。")
                return
            self.mode = "rowedit"
            self.b_mode.config(relief="sunken", text="通常表示に戻す")
            self._update_rowedit_label()
        else:
            self.mode = "normal"
            self.b_mode.config(relief="raised", text="行対応編集")
            self.left_sel_row = self.right_sel_row = None
            self.rowedit_label.config(
                text="通常モード：セルを選択すると詳細・同一値を表示します。")
        self.redraw_all()
        self._refresh_buttons()

    # ================================================================= 表示切替
    def toggle_transpose(self):
        if not self.model:
            self.transposed = not self.transposed
            return
        # 行対応編集中は通常モードへ戻す（仕様 5.2 / 4.14）
        if self.mode == "rowedit":
            self.mode = "normal"
            self.b_mode.config(relief="raised", text="行対応編集")
            self.left_sel_row = self.right_sel_row = None
        self.transposed = not self.transposed
        self.b_transpose.config(relief="sunken" if self.transposed else "raised")
        # 選択セルは (side,r,c) 不変なので維持される
        if self.single_col:
            # 向きが変わると列軸の意味が変わるため、選択から表示列を取り直す
            self.single_index = self._derive_single_index()
        self.redraw_all()
        if self.selection:
            self._reeval_same_value()
            self.redraw_all()
        self._update_single_status()
        self._refresh_buttons()

    def on_highlight_toggle(self):
        self.toggle_highlight_btn(sync_var=False)

    def toggle_highlight_btn(self, sync_var=True):
        if sync_var:
            self.highlight_enabled.set(not self.highlight_enabled.get())
        self.b_highlight.config(
            relief="sunken" if self.highlight_enabled.get() else "raised")
        self._reeval_same_value()
        self.redraw_all()

    def _reeval_same_value(self):
        if self.selection and self.mode == "normal":
            self._update_same_value(*self.selection)
        else:
            self.same_left.clear()
            self.same_right.clear()

    # ================================================================= 差分移動
    def goto_diff(self, direction: int):
        if not self.model:
            return
        pairs = self.model.pairs
        diffs = [k for k, p in enumerate(pairs) if p.status != "equal"]
        if not diffs:
            return
        cur = self._current_slot()
        if direction > 0:
            nxt = next((k for k in diffs if k > cur), diffs[0])
        else:
            nxt = next((k for k in reversed(diffs) if k < cur), diffs[-1])
        self._select_slot(nxt)

    def _current_slot(self) -> int:
        if not self.selection or not self.model:
            return -1
        side, r, c = self.selection
        for k, p in enumerate(self.model.pairs):
            dr = p.left if side == "left" else p.right
            if dr == r:
                return k
        return -1

    def _select_slot(self, k: int):
        p = self.model.pairs[k]
        if p.left is not None and p.right is not None:
            col = min(p.changed_cols) if p.changed_cols else 0
            self.on_cell_select("left", p.left, col)
        elif p.left is not None:
            self.on_cell_select("left", p.left, 0)
        else:
            self.on_cell_select("right", p.right, 0)
        self._ensure_visible(k, 0)
        self._select_difflist_for_slot(k)

    # ================================================================= Undo/Redo
    def do_undo(self):
        if self.model and self.model.can_undo():
            self.model.undo()
            self._after_map_change()

    def do_redo(self):
        if self.model and self.model.can_redo():
            self.model.redo()
            self._after_map_change()

    # ================================================================= 差分一覧
    def _rebuild_difflist(self):
        self.difflist.delete(0, tk.END)
        self._difflist_slots = []
        if not self.model:
            return
        for k, p in enumerate(self.model.pairs):
            if p.status == "equal":
                continue
            if p.status == "changed":
                cols = "、".join(self._col_label(c) for c in sorted(p.changed_cols))
                tag = "🔗変更" if p.manual else "変更"
                txt = f"{tag} 左{p.left + 1}↔右{p.right + 1}（{cols}）"
            elif p.status == "left_only":
                txt = f"行削除 左{p.left + 1}"
            else:
                txt = f"行追加 右{p.right + 1}"
            self.difflist.insert(tk.END, txt)
            self._difflist_slots.append(k)

    def _col_label(self, c: int) -> str:
        # ヘッダ行（1行目）の値を列名として使う
        if self.left_sheet and self.left_sheet.nrows > 0:
            name = self.left_sheet.text(0, c)
            if name:
                return name
        from .model import col_letter
        return col_letter(c)

    def _on_difflist_select(self, _event):
        sel = self.difflist.curselection()
        if not sel:
            return
        k = self._difflist_slots[sel[0]]
        self._select_slot(k)

    def _select_difflist_for_slot(self, k: int):
        if k in getattr(self, "_difflist_slots", []):
            i = self._difflist_slots.index(k)
            self.difflist.selection_clear(0, tk.END)
            self.difflist.selection_set(i)
            self.difflist.see(i)

    # ================================================================= コンテキスト
    def show_context_menu(self, side, dr, x_root, y_root):
        if self.mode != "rowedit":
            return
        menu = tk.Menu(self, tearoff=0)
        if side == "left":
            self.left_sel_row = dr
        else:
            self.right_sel_row = dr
        self._update_rowedit_label()
        self.redraw_all()
        menu.add_command(label="選択した左右の行を対応付ける", command=self.do_pair)
        menu.add_command(label="行対応を解除する", command=self.do_unpair)
        menu.add_command(label="自動対応に戻す（全体）", command=self.do_restore_auto)
        menu.tk_popup(x_root, y_root)

    # ================================================================= 集計/状態
    def _update_summary(self):
        if not self.model:
            return
        s = self.model.summary()
        self.status_sum.config(
            text=f"集計：追加{s['added']} 削除{s['removed']} 変更{s['changed']} "
                 f"変更セル{s['changed_cells']} 手動対応{s['manual']}")
        if self.single_col:
            self._update_single_status()
        elif s["manual"]:
            self.status_map.config(text=f"手動対応：{s['manual']}件")
        else:
            self.status_map.config(text="対応：自動のみ")

    def _refresh_buttons(self):
        has = self.model is not None
        state = "normal" if has else "disabled"
        for b in (self.b_prev, self.b_next, self.b_transpose, self.b_single,
                  self.b_highlight, self.b_mode, self.b_pair, self.b_unpair,
                  self.b_restore):
            b.config(state=state)
        col_state = "normal" if has and self.single_col else "disabled"
        self.b_colprev.config(state=col_state)
        self.b_colnext.config(state=col_state)
        self.b_undo.config(
            state="normal" if has and self.model.can_undo() else "disabled")
        self.b_redo.config(
            state="normal" if has and self.model.can_redo() else "disabled")
        rowedit = self.mode == "rowedit"
        for b in (self.b_pair, self.b_unpair):
            b.config(state="normal" if has and rowedit else "disabled")


def main(left=None, right=None):
    app = App()
    if left and right:
        try:
            app.load_pair(left, right)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("読込エラー", str(e))
    app.mainloop()


if __name__ == "__main__":
    import sys
    a = sys.argv[1:]
    main(a[0] if len(a) > 0 else None, a[1] if len(a) > 1 else None)
