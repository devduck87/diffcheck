"""3ファイル比較（A/B/C）をウィンドウ非表示で検証する。"""

import os

from exceldiff.app import App

base = os.path.join(os.path.dirname(__file__), "sample_data")
paths = [os.path.join(base, n) for n in ("left.xlsx", "right.xlsx", "third.xlsx")]

app = App()
app.withdraw()
app.load_files(*paths)
app.update_idletasks()

print("パネル数:", app.npanes, "表示パネル:", app.active_sides())
print("集計:", app.model.summary())
print("ステータス:", app.status_sum.cget("text"))

print("\n差分一覧:")
for i in range(app.difflist.size()):
    print("  ", app.difflist.get(i))

# セル選択 -> 3パネル分の詳細値
app.on_cell_select("left", 4, 2)      # A004 の価格（C だけ 1600）
app.update_idletasks()
print("\nセル詳細:", app.status_cell.cget("text"))
for s in app.active_sides():
    txt = app.detail_vals[s].get("1.0", "end").strip()
    print(f"  {s}: {txt!r}")

# セル状態（描画色の元）
k = app.model.slot_of("left", 4)
slot = app.model.pairs[k]
print("  状態:", {s: app.model.cell_state(s, slot, 2) for s in app.active_sides()})

# B にしか無い行 / A に無い行の状態
k3 = app.model.slot_of("left", 3)     # A003: B に無い
print("A003 行の状態:",
      {s: app.model.cell_state(s, app.model.pairs[k3], 0) for s in app.active_sides()})
k7 = app.model.slot_of("right", 5)    # A007: B と C だけにある追加行
print("A007 行（B/C にだけ追加）:", app.model.pairs[k7].rows,
      {s: app.model.cell_state(s, app.model.pairs[k7], 0) for s in app.active_sides()})

# 同一値ハイライト（3パネル）
app.highlight_enabled.set(True)
app.on_cell_select("left", 1, 0)      # A001
app.update_idletasks()
print("\n同一値:", app.status_same.cget("text"))
print("  A:", sorted(app.same["left"]), "B:", sorted(app.same["right"]),
      "C:", sorted(app.same["third"]))

# 差分移動
app.goto_diff(1)
print("\n次差分:", app.selection)

# 1列表示（3パネル分の行高が揃うこと）
app.on_cell_select("left", 2, 4)
app.toggle_single()
app.update_idletasks()
print("\n[1列表示]", app.single_col_header_label(), "表示行数:", len(app.single_offsets))
for s in app.active_sides():
    info = app.single_cell(s, 2)
    print(f"  {s}: {info[4]!r} ({info[3]})")
app.toggle_single()

# 行対応編集: A5 / B4 / C6 を1スロットにまとめる
app.toggle_mode()
app.on_row_select("left", 5)
app.on_row_select("right", 4)
app.on_row_select("third", 6)
print("\n", app.rowedit_label.cget("text"))
app.do_pair()
app.update_idletasks()
k = app.model.slot_of("left", 5)
print("手動対応後のスロット:", app.model.pairs[k].rows, "manual:", app.model.pairs[k].manual)
print("集計:", app.model.summary())

app.do_unpair()
print("対応解除後 集計:", app.model.summary())
app.do_undo()
app.do_undo()
print("Undo x2 後 集計:", app.model.summary())
app.toggle_mode()

# 3ファイル -> 2ファイルへ戻す
app.close_third()
app.update_idletasks()
print("\nC を閉じた後 パネル数:", app.npanes, app.active_sides())
print("集計:", app.model.summary())
print("Cグリッド表示中:", bool(app.grids["third"].winfo_manager()))

# もう一度 C を読み込む（メニュー経由と同じ処理）
app.sheets["third"] = app.wbs["third"] = None
app.load_files(*paths)
app.update_idletasks()
print("C を再読込 パネル数:", app.npanes,
      "Cグリッド表示中:", bool(app.grids["third"].winfo_manager()))

app.destroy()
print("\nTHREE-FILE SMOKE OK")
