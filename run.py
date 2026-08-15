"""起動スクリプト。

    py run.py                              # 空の状態で起動（ファイルメニューから開く）
    py run.py a.xlsx b.xlsx                # 2ファイルを指定して起動
    py run.py a.xlsx b.xlsx c.xlsx         # 3ファイル比較（基準は1つ目）
    py run.py --sample                     # 同梱サンプル2ファイルで起動
    py run.py --sample3                    # 同梱サンプル3ファイルで起動
"""

import os
import sys

from exceldiff.app import main


def _sample_paths(n: int = 2):
    base = os.path.join(os.path.dirname(__file__), "sample_data")
    names = ["left.xlsx", "right.xlsx", "third.xlsx"][:n]
    return [os.path.join(base, name) for name in names]


if __name__ == "__main__":
    args = sys.argv[1:]
    if args and args[0] in ("--sample", "--sample3"):
        paths = _sample_paths(3 if args[0] == "--sample3" else 2)
        if not all(os.path.exists(p) for p in paths):
            import make_samples
            make_samples.main()
        main(*paths)
    elif len(args) >= 2:
        main(args[0], args[1], args[2] if len(args) > 2 else None)
    else:
        main()
