"""GUI を起動せずに読込・差分・同一値・手動対応・Undo を検証する。"""

from exceldiff import reader
from exceldiff.diffengine import DiffModel
from exceldiff.normalize import NormalizeOptions
from exceldiff.valueindex import ValueIndex
from exceldiff.model import cell_address

opts = NormalizeOptions(trim=True, number_equiv=True)

lwb = reader.load("sample_data/left.xlsx")
rwb = reader.load("sample_data/right.xlsx")
left = lwb.sheets[0]
right = rwb.sheets[0]
print("読込:", left, right)

model = DiffModel(left, right, opts)

print("\n--- 自動整列 ---")
for i, p in enumerate(model.pairs):
    ltxt = left.text(p.left, 0) if p.left is not None else "   -"
    rtxt = right.text(p.right, 0) if p.right is not None else "   -"
    print(f"[{i}] L={p.left} ({ltxt:>5})  R={p.right} ({rtxt:>5})  {p.status}  changed={sorted(p.changed_cols)}")

print("\n集計:", model.summary())

print("\n--- 同一値インデックス (A001) ---")
idx = ValueIndex(left, right, opts)
lc, rc = idx.find("A001")
print("左:", [cell_address(r, c) for r, c in lc])
print("右:", [cell_address(r, c) for r, c in rc])

print("\n--- 手動対応: 左A005(行5) と 右A005改(行4) を明示対応（既に自動対応の想定）---")
# A003 削除で行位置がずれるケースを手動修正する例:
# 左のA006(行6) を 右のA006(行6) に対応させる（既に対応済みなら no-op）
before = model.summary()
model.manual_pair(5, 4)
print("手動対応後 集計:", model.summary())
print("Undo可能:", model.can_undo())
model.undo()
print("Undo後 集計:", model.summary(), "(=元:", before, ")")

print("\n--- 対応解除 -> 自動復元 ---")
model.unpair(L=1)
print("解除後:", model.summary())
model.restore_auto_all()
print("自動復元後:", model.summary())

print("\nOK")
