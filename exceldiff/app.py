"""メインアプリケーション（tkinter）。

3つの追加機能を統合する:
  * 行列入れ替え表示（左右同時）
  * 同一値セルハイライト
  * 手動での行対応修正（Undo/Redo 対応）

比較対象は2ファイル（A/B）が基本で、3つ目（C）を開くと3ファイル比較になる。
行の整列は常に A を基準とし、A↔B と A↔C の結果を1本に統合して並べる。
"""

from __future__ import annotations

import os
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk

from . import reader
from .diffengine import DiffModel
from .gridview import GridView, CW, CH, HDR_H, GRID_LINE, SINGLE_CW, SINGLE_PAD
from .model import (SIDES, SIDE_LABELS, Sheet, cell_address, col_letter,
                    side_index)
from .normalize import NormalizeOptions
from .valueindex import ValueIndex

MID_W = 66
MAX_HIGHLIGHT = 100  # 同一値ハイライトの上限（仕様 3.9）
OPEN_TYPES = [("Excel/CSV", "*.xlsx *.xlsm *.csv *.tsv *.txt"), ("すべて", "*.*")]

# 1列表示モードの行ヘッダ幅
SINGLE_HDR_W_NORMAL = 52   # 通常向き＝元の行番号
SINGLE_HDR_W_TRANS = 120   # 転置＝フィールド名を表示


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel差分比較ツール")
        self.geometry("1280x760")

        # --- 状態 -----------------------------------------------------------
        # 比較パネルは最大3つ。SIDES = ("left","right","third") をキーに持つ。
        # "third" は未読込なら None で、その場合は従来どおり2ファイル比較。
        self.wbs: dict = {s: None for s in SIDES}
        self.sheets: dict = {s: None for s in SIDES}
        self.npanes = 2               # 実際に比較しているファイル数（2 または 3）
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
        self.same: dict = {s: set() for s in SIDES}      # 同一値セル（パネル別）
        self.sel_rows: dict = {s: None for s in SIDES}   # 行対応編集の選択行
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

        # セル寸法。cell_w / cell_h は既定値。
        # 個別サイズはヘッダのドラッグで設定でき、以下に上書きを保持する。
        #   slot_ext : 整列スロット k ごとの「整列軸方向の大きさ」
        #              （全パネル共通＝行対応が揃う）
        #   field_ext: シート列 c ごとの「交差軸方向の大きさ」（パネルごとに独立）
        self.cell_w = CW
        self.cell_h = CH
        self.slot_ext: dict[int, int] = {}
        self.field_ext = {s: {} for s in SIDES}
        # 左右のスクロール位置を（整列軸に加えて）両軸で一致させるか
        self.sync_both = tk.BooleanVar(value=False)

        self._build_menu()
        self._build_toolbar()
        self._build_detail_panel()
        self._build_body()
        self._build_rowedit_panel()
        self._build_statusbar()

        self._refresh_buttons()

    # ============================================================= パネル共通
    def active_sides(self) -> list[str]:
        """比較中のパネルキー一覧（2ファイルなら left/right の2つ）。"""
        return list(SIDES[:self.npanes])

    def sheet_of(self, side) -> Sheet | None:
        return self.sheets[SIDES[side_index(side)]]

    def active_sheets(self) -> list[Sheet]:
        return [self.sheets[s] for s in self.active_sides()]

    def base_grid(self) -> GridView:
        """基準パネル（A）のグリッド。スクロール基準に使う。"""
        return self.grids["left"]

    def max_ncols(self) -> int:
        return max((sh.ncols for sh in self.active_sheets() if sh), default=0)

    # 2ファイル時代からの別名（外部スクリプト互換）
    @property
    def left_sheet(self):
        return self.sheets["left"]

    @property
    def right_sheet(self):
        return self.sheets["right"]

    @property
    def third_sheet(self):
        return self.sheets["third"]

    @property
    def same_left(self):
        return self.same["left"]

    @property
    def same_right(self):
        return self.same["right"]

    @property
    def left_sel_row(self):
        return self.sel_rows["left"]

    @property
    def right_sel_row(self):
        return self.sel_rows["right"]

    @property
    def left_grid(self):
        return self.grids["left"]

    @property
    def right_grid(self):
        return self.grids["right"]

    # ================================================================= UI構築
    def _build_menu(self):
        m = tk.Menu(self)
        self.config(menu=m)

        fm = tk.Menu(m, tearoff=0)
        fm.add_command(label="ファイルA（基準）を開く...",
                       command=lambda: self.open_file("left"))
        fm.add_command(label="ファイルB を開く...",
                       command=lambda: self.open_file("right"))
        fm.add_command(label="ファイルC を開く...（3ファイル比較）",
                       command=lambda: self.open_file("third"))
        fm.add_separator()
        fm.add_command(label="ファイルC を閉じる（2ファイル比較に戻す）",
                       command=self.close_third)
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
        vm.add_checkbutton(label="各パネルのスクロール位置を一致させる",
                           variable=self.sync_both, command=self._on_sync_toggle)
        vm.add_separator()
        vm.add_command(label="セルの大きさを既定に戻す", command=self.reset_cell_sizes)
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
        tk.Frame(bar, width=2, bg="#bbb").pack(side="left", fill="y", padx=4)
        self.b_third = btn("＋Cファイル", self.toggle_third)

    def _build_detail_panel(self):
        """各グリッドの上に、選択セルと同じ整列行の値を並べて表示する帯。

        値が長い場合に備え、各パネルの値はスクロール可能な Text で表示する。
        パネルCの枠は3ファイル比較のときだけ表示する（配置順は最後）。
        """
        p = tk.Frame(self, bd=1, relief="sunken", height=110)
        p.pack(side="top", fill="x")
        p.pack_propagate(False)
        self._detail_bar = p

        # 差分一覧の上：見出しと選択セル情報（スクロール可能）
        info = tk.Frame(p, width=230)
        info.pack(side="left", fill="y")
        info.pack_propagate(False)
        tk.Label(info, text="セル詳細", bg="#dfe4ea", anchor="w").pack(fill="x")
        self.detail_addr = self._make_scroll_text(info)

        # 各グリッドの上：そのパネルの値（全文・スクロール可能）
        self.detail_vals: dict = {}
        self._detail_frames: dict = {}
        self._detail_spacers: dict = {}
        for s in SIDES:
            if s != "left":
                # 対応列（中央キャンバス）の上のスペーサ
                sp = tk.Frame(p, width=MID_W)
                sp.pack(side="left", fill="y")
                self._detail_spacers[s] = sp
            f = tk.Frame(p)
            f.pack(side="left", fill="both", expand=True)
            tk.Label(f, text=f"{SIDE_LABELS[s]} の値", bg="#eef1f5",
                     anchor="w").pack(fill="x")
            self.detail_vals[s] = self._make_scroll_text(f)
            self._detail_frames[s] = f

        # 2ファイル比較の間はパネルCの枠を隠す
        self._detail_frames["third"].pack_forget()
        self._detail_spacers["third"].pack_forget()

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
        """選択セルと同じ整列行の各パネルの値を詳細帯に表示する。"""
        if not self.model:
            return
        k = self.model.slot_of(side, r)
        slot = self.model.pairs[k] if k is not None else None
        for s in self.active_sides():
            sheet = self.sheets[s]
            dr = slot.row(s) if slot is not None else None
            val = (sheet.text(dr, c)
                   if dr is not None and sheet and c < sheet.ncols else "")
            self._set_text(self.detail_vals[s], val)
        self._set_text(self.detail_addr,
                       f"列：{self._col_label(c)}\n"
                       f"選択：{SIDE_LABELS[side]} {cell_address(r, c)}")

    def _clear_detail(self):
        if hasattr(self, "detail_addr"):
            self._set_text(self.detail_addr, "セル未選択")
            for s in SIDES:
                self._set_text(self.detail_vals[s], "")

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

    # ---------------------------------------------------------- 個別セルサイズ
    def aligned_default(self) -> int:
        """整列軸（スロット）方向の既定サイズ。"""
        return self.cell_w if self.transposed else self.cell_h

    def cross_default(self) -> int:
        """交差軸（シート列）方向の既定サイズ。"""
        return self.cell_h if self.transposed else self.cell_w

    def slot_extent(self, k: int) -> int:
        return self.slot_ext.get(k, self.aligned_default())

    def field_extent(self, side: str, c: int) -> int:
        return self.field_ext[side].get(c, self.cross_default())

    def set_extent(self, kind: str, side: str, key: int, size: int):
        """ドラッグ結果を反映する。slot は全パネル共通、field は side ごと。"""
        size = max(14, min(600, int(size)))
        if kind == "slot":
            self.slot_ext[key] = size          # 全パネル共通 → 行対応が揃う
        else:
            self.field_ext[side][key] = size
        self.redraw_all()

    def autofit(self, kind: str, side: str, key: int):
        """境界ダブルクリックで内容に合わせて自動調整。"""
        if kind == "field":
            sheet = self.sheets[side]
            maxw = 0
            for r in range(sheet.nrows):
                t = sheet.text(r, key).replace("\n", " ")
                if t:
                    maxw = max(maxw, self._font.measure(t))
            self.field_ext[side][key] = max(40, min(600, maxw + 14))
        else:
            # 行高は既定（1行）に戻す
            self.slot_ext.pop(key, None)
        self.redraw_all()

    def reset_cell_sizes(self):
        self.slot_ext.clear()
        for s in SIDES:
            self.field_ext[s].clear()
        self.redraw_all()

    def normal_slot_offsets(self):
        """通常向きのスロット y 開始位置一覧と総高を返す。"""
        offs = []
        y = HDR_H
        for k in range(len(self.model.pairs)):
            offs.append(y)
            y += self.slot_extent(k)
        return offs, y

    def _on_sync_toggle(self):
        # 有効化した瞬間に、基準パネル(A)の現在位置へ他パネル（と中央）を合わせる
        if self.sync_both.get():
            g = self.base_grid()
            self.propagate_scroll("h", g.canvas.xview()[0], g)
            self.propagate_scroll("v", g.canvas.yview()[0], g)

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

        # グリッドと、その間の対応関係キャンバスを A|B|C の順に並べる。
        # パネルCとその手前のキャンバスは配置順の最後なので、3ファイル比較に
        # なった時点で pack すれば右端に追加される。
        self.grids: dict = {}
        self.mids: dict = {}   # キー: 右側パネル（"right" は A|B、"third" は B|C）
        for s in SIDES:
            if s != "left":
                mid = tk.Canvas(body, width=MID_W, background="#f7f8fa",
                                highlightthickness=0)
                mid.pack(side="left", fill="y")
                self.mids[s] = mid
            g = GridView(body, self, s)
            g.pack(side="left", fill="both", expand=True)
            self.grids[s] = g
        self.grids["third"].pack_forget()
        self.mids["third"].pack_forget()
        # 2ファイル時代の別名
        self.mid = self.mids["right"]

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
        self.status_same = tk.Label(sb, text="同一値：-", width=32, anchor="w")
        self.status_same.pack(side="left", padx=4)
        self.status_map = tk.Label(sb, text="対応：-", width=26, anchor="w")
        self.status_map.pack(side="left", padx=4)
        self.status_sum = tk.Label(sb, text="集計：-", anchor="w")
        self.status_sum.pack(side="left", padx=4)

    # ================================================================= ファイル
    def open_file(self, side: str):
        path = filedialog.askopenfilename(
            title=f"ファイル{SIDE_LABELS[side]} を開く", filetypes=OPEN_TYPES)
        if not path:
            return
        try:
            wb = reader.load(path)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("読込エラー", f"{os.path.basename(path)}\n{e}")
            return
        self.wbs[side] = wb
        self.sheets[side] = wb.sheets[0]
        if self.sheets["left"] and self.sheets["right"]:
            self.recompare()

    def toggle_third(self):
        """3ファイル比較の追加／解除をトグルする（ツールバー用）。"""
        if self.sheets["third"]:
            self.close_third()
        else:
            self.open_file("third")

    def close_third(self):
        """3つ目のファイルを外し、2ファイル比較へ戻す。"""
        if not self.sheets["third"]:
            return
        self.wbs["third"] = None
        self.sheets["third"] = None
        if self.sheets["left"] and self.sheets["right"]:
            self.recompare()
        else:
            self.npanes = 2
            self._apply_pane_visibility()

    def load_pair(self, left_path: str, right_path: str, third_path: str = None):
        """A/B（必要なら C）を読み込んで比較する。"""
        paths = [left_path, right_path] + ([third_path] if third_path else [])
        self.load_files(*paths)

    def load_files(self, *paths: str):
        if not 2 <= len(paths) <= len(SIDES):
            raise ValueError(f"比較できるのは2〜{len(SIDES)}ファイルです。")
        for s in SIDES:
            self.wbs[s] = None
            self.sheets[s] = None
        for s, path in zip(SIDES, paths):
            self.wbs[s] = reader.load(path)
            self.sheets[s] = self.wbs[s].sheets[0]
        self.recompare()

    def _apply_pane_visibility(self):
        """パネルC（グリッド・対応列・詳細帯）の表示/非表示を切り替える。"""
        show = self.npanes >= 3
        packed = bool(self.grids["third"].winfo_manager())
        if show and not packed:
            self.mids["third"].pack(side="left", fill="y")
            self.grids["third"].pack(side="left", fill="both", expand=True)
            self._detail_spacers["third"].pack(side="left", fill="y")
            self._detail_frames["third"].pack(side="left", fill="both",
                                              expand=True)
        elif not show and packed:
            self.mids["third"].pack_forget()
            self.grids["third"].pack_forget()
            self._detail_spacers["third"].pack_forget()
            self._detail_frames["third"].pack_forget()

    def recompare(self):
        if not (self.sheets["left"] and self.sheets["right"]):
            return
        for k, var in self._num_opts.items():
            setattr(self.opts, k, var.get())
        self.npanes = 3 if self.sheets["third"] else 2
        self._apply_pane_visibility()
        sheets = self.active_sheets()
        self.model = DiffModel(sheets, self.opts)
        self.vindex = ValueIndex(sheets, self.opts)
        # 個別サイズはファイル索引に依存するため作り直し時にクリア
        self.slot_ext.clear()
        for s in SIDES:
            self.field_ext[s].clear()
            self.same[s].clear()
            self.sel_rows[s] = None
        self.selection = None
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
        for s in self.active_sides():
            self.grids[s].redraw()
        self._draw_mid()

    def _draw_mid(self):
        """各パネルの間の対応関係キャンバスを描く。"""
        for s in self.active_sides()[1:]:
            self._draw_mid_canvas(self.mids[s], side_index(s))

    def _draw_mid_canvas(self, c, right_pane: int):
        """right_pane と その左隣（right_pane - 1）の対応を描く。"""
        c.delete("all")
        if not self.model:
            return
        left_pane = right_pane - 1
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
                self._draw_mid_connector(c, p, y, left_pane, right_pane)
            return
        offs, total = self.normal_slot_offsets()
        c.configure(scrollregion=(0, 0, MID_W, max(total, 1)))
        if self.transposed:
            c.create_text(MID_W / 2, 40, text="行列\n入替中", font=("Meiryo", 8),
                          fill="#888", justify="center")
            return
        if not self.show_lines.get():
            return
        for k, p in enumerate(self.model.pairs):
            y = offs[k] + self.slot_extent(k) / 2
            self._draw_mid_connector(c, p, y, left_pane, right_pane)

    def _draw_mid_connector(self, c, p, y, left_pane=0, right_pane=1):
        lrow, rrow = p.row(left_pane), p.row(right_pane)
        if lrow is not None and rrow is not None:
            if p.manual:
                c.create_line(4, y, MID_W - 4, y, fill="#1a73e8", width=3)
                c.create_text(MID_W / 2, y - 7, text="🔗", font=("Meiryo", 7))
            else:
                c.create_line(6, y, MID_W - 6, y, fill="#9aa0a6", width=1)
        elif lrow is not None:
            c.create_text(MID_W / 2, y, text="◀削除", font=("Meiryo", 7),
                          fill="#c0392b")
        elif rrow is not None:
            c.create_text(MID_W / 2, y, text="追加▶", font=("Meiryo", 7),
                          fill="#1e824c")

    # ---------------------------------------------------------- 1列表示レイアウト
    def _single_ndisp(self) -> int:
        """1列表示の表示行数（現在の向きの行軸）。"""
        if self.transposed:
            return self.max_ncols()
        return len(self.model.pairs)

    def single_cols_count(self) -> int:
        """1列表示で選択できる列の総数（現在の向きの列軸）。"""
        if self.transposed:
            return len(self.model.pairs)
        return self.max_ncols()

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
        dr = pair.row(side)
        state = self.model.cell_state(side, pair, col)
        sheet = self.sheets[side]
        text = (sheet.text(dr, col)
                if (dr is not None and sheet and col < sheet.ncols) else "")
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
        """全パネル共通の可変行高を計算し、対応行を揃える。"""
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
        sides = self.active_sides()
        for i in range(nd):
            lines = 1
            for s in sides:
                info = self.single_cell(s, i)
                if info:
                    lines = max(lines, self._wrap_lines(info[4]))
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
                parts = []
                for s in self.active_sides():
                    dr = p.row(s)
                    parts.append(f"{SIDE_LABELS[s]}"
                                 f"{dr + 1 if dr is not None else '-'}")
                return " / ".join(parts) + " 行"
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
        # 行＝整列スロット（元の行番号）。パネルごとに異なるので side ごとに
        # 出す方が正確だが、簡易表示として存在する先頭パネルの行番号を出す。
        if 0 <= i < len(self.model.pairs):
            p = self.model.pairs[i]
            present = p.present()
            if present:
                return str(p.rows[present[0]] + 1)
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
                k = self.model.slot_of(side, r)
                if k is not None:
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
            sides = self.active_sides()
            targets = [self.grids[s].canvas for s in sides]
            if axis == "v":
                targets.extend(self.mids[s] for s in sides[1:])
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
        base = self.base_grid()
        if self.single_col:
            # 表示行 i を割り出して縦位置へスクロール
            i = k if not self.transposed else col
            if 0 <= i < len(self.single_offsets) and self.single_total > 0:
                frac = max(0, self.single_offsets[i] / self.single_total - 0.05)
                base.canvas.yview_moveto(frac)
                self.propagate_scroll("v", frac, base)
            return
        if self.transposed:
            # 縦＝シート列、横＝スロット。基準パネルの寸法で概算スクロールする。
            ysum = HDR_H
            for cc in range(col):
                ysum += self.field_extent("left", cc)
            ytotal = HDR_H
            for cc in range(self.left_sheet.ncols):
                ytotal += self.field_extent("left", cc)
            frac_y = max(0, ysum / max(ytotal, 1) - 0.1)
            for s in self.active_sides():
                self.grids[s].canvas.yview_moveto(frac_y)
            xsum = HDR_W
            for kk in range(k):
                xsum += self.slot_extent(kk)
            xtotal = HDR_W + sum(self.slot_extent(kk)
                                 for kk in range(len(self.model.pairs)))
            frac = max(0, xsum / max(xtotal, 1) - 0.1)
            base.canvas.xview_moveto(frac)
            self.propagate_scroll("h", frac, base)
        else:
            offs, total = self.normal_slot_offsets()
            frac = max(0, offs[k] / max(total, 1) - 0.1)
            base.canvas.yview_moveto(frac)
            self.propagate_scroll("v", frac, base)

    # ================================================================= 選択
    def on_cell_select(self, side: str, r: int, c: int):
        self.selection = (side, r, c)
        self._update_same_value(side, r, c)
        self._update_cell_status(side, r, c)
        self._update_detail(side, r, c)
        self.redraw_all()

    def _update_same_value(self, side, r, c):
        for s in SIDES:
            self.same[s].clear()
        self._same_nav = []
        self._same_nav_i = -1
        if not self.highlight_enabled.get() or not self.vindex:
            self.status_same.config(text="同一値：-")
            return
        sheet = self.sheets[side]
        cell = sheet.cell(r, c)
        key = self.model.cell_key(side, r, c)
        if cell.is_blank() or key == "":
            self.status_same.config(text="同一値：空白は対象外")
            return
        sides = self.active_sides()
        found = self.vindex.find(key)
        # 表示上限（仕様 3.9）。先頭のパネルから順に埋める。
        capped = ""
        if sum(len(f) for f in found) > MAX_HIGHLIGHT:
            keep = MAX_HIGHLIGHT
            trimmed = []
            for cells in found:
                take = cells[:max(0, keep)]
                keep -= len(take)
                trimmed.append(take)
            found = trimmed
            capped = f"（先頭{MAX_HIGHLIGHT}件を表示）"
        for s, cells in zip(sides, found):
            self.same[s] = set(cells)
            self._same_nav.extend((s, *p) for p in cells)
        counts = "　".join(f"{SIDE_LABELS[s]}{len(f)}件"
                           for s, f in zip(sides, found))
        self.status_same.config(text=f"同一値：{counts}{capped}")

    def _update_cell_status(self, side, r, c):
        self.status_cell.config(
            text=f"元セル：{SIDE_LABELS[side]} {cell_address(r, c)}")

    # ================================================================= 行対応編集
    def on_row_select(self, side: str, dr: int):
        self.sel_rows[side] = dr
        self._update_rowedit_label()
        self.redraw_all()

    def _selected_rows(self) -> list[tuple[str, int]]:
        """行対応編集で選択中の (パネル, 行) 一覧（表示順）。"""
        return [(s, self.sel_rows[s]) for s in self.active_sides()
                if self.sel_rows[s] is not None]

    def _update_rowedit_label(self):
        parts = []
        for s in self.active_sides():
            dr = self.sel_rows[s]
            parts.append(f"{SIDE_LABELS[s]}選択行："
                         f"{dr + 1 if dr is not None else '-'}"
                         + ("行目" if dr is not None else ""))
        self.rowedit_label.config(
            text="行対応編集モード｜" + "　".join(parts)
                 + "　→［対応付け］で同じ行として対応させます")

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
        sel = self._selected_rows()
        if len(sel) < 2:
            messagebox.showwarning(
                "選択不足",
                "行が2つ以上選択されていません。\n"
                "対応させたいパネルから1行ずつ選択してください。")
            return
        anchor_side, anchor_row = sel[0]
        targets = [(s, r) for s, r in sel[1:]
                   if not self.model.is_paired(anchor_side, anchor_row, s, r)]
        if not targets:
            messagebox.showinfo("行対応", "選択した行はすでに対応付けられています。")
            return
        for s, r in targets:
            partner = self.model.partner_of(s, r, anchor_side)
            if partner is not None and partner != anchor_row:
                ok = messagebox.askyesno(
                    "対応の競合",
                    f"{SIDE_LABELS[s]}{r + 1}行目は、すでに "
                    f"{SIDE_LABELS[anchor_side]}{partner + 1}行目と"
                    f"対応付けられています。\n\n"
                    f"既存の対応を解除して、"
                    f"{SIDE_LABELS[anchor_side]}{anchor_row + 1}行目と"
                    f"対応付けますか？")
                if not ok:
                    continue
            self.model.manual_pair(anchor_side, anchor_row, s, r)
        self._after_map_change()

    def do_unpair(self):
        if not self.model:
            return
        if self.mode != "rowedit":
            messagebox.showinfo("行対応編集", "「行対応編集」モードに切り替えてください。")
            return
        sel = self._selected_rows()
        if not sel:
            messagebox.showwarning("選択不足", "解除する行を選択してください。")
            return
        side, row = sel[0]
        self.model.unpair(side, row)
        self._after_map_change()

    def do_restore_auto(self):
        if not self.model:
            return
        if messagebox.askyesno("自動対応に戻す",
                               "手動対応をすべて破棄し、自動対応へ戻しますか？"):
            self.model.restore_auto_all()
            self._after_map_change()

    def _clear_row_sel(self):
        for s in SIDES:
            self.sel_rows[s] = None

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
            self._clear_row_sel()
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
            self._clear_row_sel()
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
            for s in SIDES:
                self.same[s].clear()

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
        k = self.model.slot_of(side, r)
        return -1 if k is None else k

    def _select_slot(self, k: int):
        """スロット k の代表セル（存在する先頭パネルの変更列）を選択する。"""
        p = self.model.pairs[k]
        present = p.present()
        if not present:
            return
        i = present[0]
        col = min(p.changed_cols) if p.changed_cols else 0
        self.on_cell_select(SIDES[i], p.rows[i], col)
        self._ensure_visible(k, col if self.transposed else 0)
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
        sides = self.active_sides()
        for k, p in enumerate(self.model.pairs):
            if p.status == "equal":
                continue
            # 各パネルの行番号（無いパネルは "-"）
            where = "↔".join(
                f"{SIDE_LABELS[s]}"
                f"{p.row(s) + 1 if p.row(s) is not None else '-'}"
                for s in sides)
            if p.changed_cols:
                cols = "、".join(self._col_label(c) for c in sorted(p.changed_cols))
                tag = "🔗変更" if p.manual else "変更"
                txt = f"{tag} {where}（{cols}）"
            elif p.rows[0] is None:
                txt = f"行追加 {where}"
            else:
                txt = f"行削除 {where}"
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
        self.sel_rows[side] = dr
        self._update_rowedit_label()
        self.redraw_all()
        menu.add_command(label="選択した行どうしを対応付ける", command=self.do_pair)
        menu.add_command(label="行対応を解除する", command=self.do_unpair)
        menu.add_command(label="自動対応に戻す（全体）", command=self.do_restore_auto)
        menu.tk_popup(x_root, y_root)

    # ================================================================= 集計/状態
    def _update_summary(self):
        if not self.model:
            return
        s = self.model.summary()
        files = f"{self.npanes}ファイル比較"
        self.status_sum.config(
            text=f"集計（{files}）：追加{s['added']} 削除{s['removed']} "
                 f"変更{s['changed']} 変更セル{s['changed_cells']} "
                 f"手動対応{s['manual']}")
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
        self.b_third.config(
            text="Cを閉じる" if self.sheets["third"] else "＋Cファイル")


def main(left=None, right=None, third=None):
    app = App()
    if left and right:
        try:
            app.load_pair(left, right, third)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("読込エラー", str(e))
    app.mainloop()


if __name__ == "__main__":
    import sys
    a = sys.argv[1:]
    main(a[0] if len(a) > 0 else None, a[1] if len(a) > 1 else None,
         a[2] if len(a) > 2 else None)
