"""個別セルサイズ（列幅・行高）機能をウィンドウ非表示で検証する。"""

import os
from exceldiff.app import App
from exceldiff.gridview import HDR_W, HDR_H

base = os.path.join(os.path.dirname(__file__), "sample_data")
app = App()
app.withdraw()
app.load_pair(os.path.join(base, "left.xlsx"), os.path.join(base, "right.xlsx"))
app.update_idletasks()

g = app.left_grid
g.redraw()
print("既定の横境界(先頭5):", g._hx[:5])
print("既定の縦境界(先頭5):", g._vy[:5])

# 列0(左)の幅を200へ、スロット2の行高を60へ設定
app.set_extent("field", "left", 0, 200)
app.set_extent("slot", "left", 2, 60)
g.redraw()
print("\n列0幅を200に:", g._hx[1] - g._hx[0], "==200?")
print("スロット2高を60に:", g._vy[3] - g._vy[2], "==60?")

# 行高(slot)は左右共通で揃うことを確認
app.right_grid.redraw()
rg = app.right_grid
print("右側スロット2高:", rg._vy[3] - rg._vy[2], "(左右一致で行対応が揃う)")

# 列幅は左右独立（右の列0は既定のまま）
print("右側 列0幅:", rg._hx[1] - rg._hx[0], "(=既定", app.cell_w, ")")

# クリック座標→セルのヒットテスト（可変サイズでも正しく当たるか）
class E:  # 疑似イベント
    def __init__(self, x, y):
        self.x, self.y = x, y
# 列0(幅200)の中央、スロット0付近をクリック
hit = g._hit(E(HDR_W + 100, HDR_H + 5))
print("\nヒットテスト (列0内, スロット0):", hit, "-> col=0期待")
# 列1(既定幅)の中央
hit2 = g._hit(E(HDR_W + 200 + app.cell_w // 2, HDR_H + 5))
print("ヒットテスト (列1内):", hit2, "-> col=1期待")

# 境界検出（列0の右端 x=HDR_W+200 付近, 上ヘッダ内 cy<HDR_H）
div = g._divider_at(HDR_W + 200, HDR_H - 3)
print("\n境界検出 上ヘッダ 列0右端:", div, "-> ('H',0)期待")
div2 = g._divider_at(3, HDR_H + 60)  # スロット0(既定)下端は HDR_H+cell_h
print("境界検出 左ヘッダ:", div2)

# 自動フィット（列0幅を内容に合わせる）
app.autofit("field", "left", 0)
g.redraw()
print("\n列0 自動フィット後の幅:", g._hx[1] - g._hx[0])

# リセット
app.reset_cell_sizes()
g.redraw()
print("リセット後 列0幅:", g._hx[1] - g._hx[0], "スロット2高:", g._vy[3] - g._vy[2])

# 転置しても破綻しないか
app.toggle_transpose()
app.set_extent("slot", "left", 1, 150)  # 転置時: スロットは横軸(幅)
g.redraw()
print("\n[転置] スロット1の横幅:", g._hx[2] - g._hx[1], "==150?")

app.destroy()
print("\nSIZE SMOKE OK")
