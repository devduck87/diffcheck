"""値の正規化（一致判定条件）。

仕様 3.5 / 3.6 に対応。差分比較と同一値ハイライトで正規化条件を共用する（初期仕様）。
"""

from __future__ import annotations

import unicodedata


class NormalizeOptions:
    """一致判定の正規化オプション。"""

    def __init__(
        self,
        ignore_case: bool = False,
        trim: bool = False,
        fullwidth: bool = False,      # 全角・半角を同一視（NFKC）
        number_equiv: bool = False,   # 数値と数値文字列を同一視
        blank_equiv: bool = True,     # 空白セルと空文字列を同一視
    ):
        self.ignore_case = ignore_case
        self.trim = trim
        self.fullwidth = fullwidth
        self.number_equiv = number_equiv
        self.blank_equiv = blank_equiv


def norm_key(text: str, opts: NormalizeOptions) -> str:
    """正規化済みの比較キー文字列を返す。空相当なら "" を返す。"""
    s = text if text is not None else ""
    if opts.fullwidth:
        s = unicodedata.normalize("NFKC", s)
    if opts.trim:
        s = s.strip()
    if opts.blank_equiv and s.strip() == "":
        return ""
    if opts.ignore_case:
        s = s.casefold()
    if opts.number_equiv:
        s = _canonical_number(s)
    return s


def _canonical_number(s: str) -> str:
    try:
        f = float(s)
    except (ValueError, TypeError):
        return s
    if f.is_integer():
        return str(int(f))
    return repr(f)
