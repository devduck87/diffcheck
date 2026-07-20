"""Excel差分比較ツール（標準ライブラリのみ）。

追加機能仕様書に基づく実装:
  * 行列入れ替え表示
  * 同一値セルのハイライト
  * 手動での行対応修正
"""

__all__ = [
    "model",
    "normalize",
    "reader",
    "diffengine",
    "valueindex",
]
