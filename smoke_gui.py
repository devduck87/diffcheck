"""GUI をウィンドウ非表示(withdraw)で構築し、主要操作を無人で検証する。"""

import os
from exceldiff.app import App

base = os.path.join(os.path.dirname(__file__), "sample_data")
left = os.path.join(base, "left.xlsx")
right = os.path.join(base, "right.xlsx")

app = App()
app.withdraw()               # 画面に出さない
app.load_pair(left, right)
app.update_idletasks()
print("差分一覧件数:", app.difflist.size())
print("集計:", app.model.summary())

# 選択 + 同一値ハイライト
app.highlight_enabled.set(True)
app.on_cell_select("left", 1, 0)   # A001
app.update_idletasks()
print("同一値 左:", sorted(app.same_left), "右:", sorted(app.same_right))
print("status_same:", app.status_same.cget("text"))

# 差分移動
app.goto_diff(1)
app.update_idletasks()
print("次差分 選択:", app.selection)

# 行列入替（選択維持を確認）
sel_before = app.selection
app.toggle_transpose()
app.update_idletasks()
print("転置後 selection(維持):", app.selection, "==", sel_before)
app.toggle_transpose()  # 戻す

# 行対応編集: L3(削除) と R5(追加) を手動対応
app.toggle_mode()
app.on_row_select("left", 3)
app.on_row_select("right", 5)
app.do_pair()
app.update_idletasks()
print("手動対応後 集計:", app.model.summary())

# Undo / Redo
app.do_undo()
print("Undo後 集計:", app.model.summary())
app.do_redo()
print("Redo後 集計:", app.model.summary())

# --- 1列表示モード -----------------------------------------------------------
app.toggle_mode()  # rowedit 解除
app.on_cell_select("left", 2, 2)   # A002 の価格セル（変更あり）を選択
app.toggle_single()
app.update_idletasks()
print("\n[1列表示 通常向き]")
print("  選択から表示列:", app.single_index, "ラベル:", app.single_col_header_label())
print("  表示行数:", len(app.single_offsets), "総高:", app.single_total)
print("  status:", app.status_map.cget("text"))
# 列送り
app.step_column(1)
print("  列送り後:", app.single_index, app.single_col_header_label())

# 転置 + 1列表示（1レコードをフィールド単位で縦比較）
app.toggle_transpose()
app.update_idletasks()
print("\n[1列表示 + 転置（1レコード比較）]")
print("  表示列(レコード):", app.single_index, "->", app.single_col_header_label())
print("  表示行数(フィールド):", len(app.single_offsets))
for i in range(min(3, len(app.single_offsets))):
    li = app.single_cell("left", i)
    ri = app.single_cell("right", i)
    print(f"   行{i} {app.single_row_header_label(i)}: 左={li[4]!r} 右={ri[4]!r} 状態={li[3]}/{ri[3]}")

# 長文の折り返し高さ確認
long_left = app.left_sheet.rows[1][4]
long_left.value = long_left.display = "とても長い名前が入っているセルの内容確認用テキスト" * 3
app.build_single_layout()
print("\n長文セルの折り返し行数:", app._wrap_lines(long_left.display))

app.destroy()
print("\nSMOKE OK")
