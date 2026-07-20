"""起動スクリプト。

    py run.py                         # 空の状態で起動（ファイルメニューから開く）
    py run.py left.xlsx right.xlsx    # 2ファイルを指定して起動
    py run.py --sample                # 同梱サンプルで起動
"""

import os
import sys

from exceldiff.app import main


def _sample_paths():
    base = os.path.join(os.path.dirname(__file__), "sample_data")
    return (os.path.join(base, "left.xlsx"), os.path.join(base, "right.xlsx"))


if __name__ == "__main__":
    args = sys.argv[1:]
    if args and args[0] == "--sample":
        left, right = _sample_paths()
        if not (os.path.exists(left) and os.path.exists(right)):
            import make_samples
            make_samples.main()
        main(left, right)
    elif len(args) >= 2:
        main(args[0], args[1])
    else:
        main()
