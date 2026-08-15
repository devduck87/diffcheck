"""GUI を起動せずに読込・差分・同一値・手動対応・Undo を検証する。"""

from exceldiff import reader
from exceldiff.diffengine import DiffModel
from exceldiff.model import SIDE_LABELS, SIDES, cell_address
from exceldiff.normalize import NormalizeOptions
from exceldiff.valueindex import ValueIndex

opts = NormalizeOptions(trim=True, number_equiv=True)

lwb = reader.load("sample_data/left.xlsx")
rwb = reader.load("sample_data/right.xlsx")
left = lwb.sheets[0]
right = rwb.sheets[0]
print("読込:", left, right)

model = DiffModel([left, right], opts)

print("\n--- 自動整列 ---")
for i, p in enumerate(model.pairs):
    ltxt = left.text(p.left, 0) if p.left is not None else "   -"
    rtxt = right.text(p.right, 0) if p.right is not None else "   -"
    print(f"[{i}] L={p.left} ({ltxt:>5})  R={p.right} ({rtxt:>5})  {p.status}  changed={sorted(p.changed_cols)}")

print("\n集計:", model.summary())

print("\n--- 同一値インデックス (A001) ---")
idx = ValueIndex([left, right], opts)
lc, rc = idx.find("A001")
print("左:", [cell_address(r, c) for r, c in lc])
print("右:", [cell_address(r, c) for r, c in rc])

print("\n--- 手動対応: 左A005(行5) と 右A005改(行4) を明示対応（既に自動対応の想定）---")
# A003 削除で行位置がずれるケースを手動修正する例:
# 左のA006(行6) を 右のA006(行6) に対応させる（既に対応済みなら no-op）
before = model.summary()
model.manual_pair("left", 5, "right", 4)
print("手動対応後 集計:", model.summary())
print("Undo可能:", model.can_undo())
model.undo()
print("Undo後 集計:", model.summary(), "(=元:", before, ")")

print("\n--- 対応解除 -> 自動復元 ---")
model.unpair("left", 1)
print("解除後:", model.summary())
model.restore_auto_all()
print("自動復元後:", model.summary())

# --------------------------------------------------------------- 3ファイル比較
print("\n=== 3ファイル比較（基準=A） ===")
third = reader.load("sample_data/third.xlsx").sheets[0]
m3 = DiffModel([left, right, third], opts)
sheets = {"left": left, "right": right, "third": third}
print("パネル数:", m3.npanes)
for i, p in enumerate(m3.pairs):
    cells = []
    for s in SIDES:
        dr = p.row(s)
        txt = sheets[s].text(dr, 0) if dr is not None else "-"
        cells.append(f"{SIDE_LABELS[s]}={dr if dr is not None else '-'}({txt})")
    print(f"[{i}] {'  '.join(cells)}  {p.status}  changed={sorted(p.changed_cols)}")
print("集計:", m3.summary())

print("\n--- セル状態（A004 の価格列: A/B は 1500、C は 1600）---")
for i, p in enumerate(m3.pairs):
    if p.row("left") is not None and left.text(p.row("left"), 0) == "A004":
        for s in SIDES:
            print(f"  {SIDE_LABELS[s]}: {m3.cell_state(s, p, 2)}")
        break

print("\n--- 3ファイルでの手動対応 / 解除 ---")
before3 = m3.summary()
m3.manual_pair("left", 5, "third", 6)   # A005 に C の A007 行をぶつける
print("手動対応後:", m3.summary())
m3.undo()
print("Undo後:", m3.summary(), "(=元:", before3, ")")
m3.unpair("left", 2)
print("A行2の対応解除後:", m3.summary())
m3.restore_auto_all()
print("自動復元後:", m3.summary())

print("\nOK")
