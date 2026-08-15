"""動作確認用のサンプル xlsx を標準ライブラリのみで生成する。

openpyxl を使わず、最小限の xlsx（zip+XML）を書き出す。
"""

from __future__ import annotations

import os
import zipfile

from exceldiff.model import col_letter

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

WB_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"""


def _xml_escape(s: str) -> str:
    return (s.replace("&", "&amp;").replace("<", "&lt;")
             .replace(">", "&gt;").replace('"', "&quot;"))


def write_xlsx(path: str, rows: list[list], sheet_name: str = "Sheet1"):
    shared: list[str] = []
    shared_idx: dict[str, int] = {}

    def sref(text: str) -> int:
        if text not in shared_idx:
            shared_idx[text] = len(shared)
            shared.append(text)
        return shared_idx[text]

    # sheet XML
    row_xml = []
    for ri, row in enumerate(rows, start=1):
        cells = []
        for ci, val in enumerate(row):
            ref = f"{col_letter(ci)}{ri}"
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                cells.append(f'<c r="{ref}"><v>{val}</v></c>')
            else:
                idx = sref("" if val is None else str(val))
                cells.append(f'<c r="{ref}" t="s"><v>{idx}</v></c>')
        row_xml.append(f'<row r="{ri}">{"".join(cells)}</row>')

    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetData>{"".join(row_xml)}</sheetData></worksheet>'
    )

    sst_items = "".join(f"<si><t>{_xml_escape(s)}</t></si>" for s in shared)
    sst_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        f'count="{len(shared)}" uniqueCount="{len(shared)}">{sst_items}</sst>'
    )

    wb_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<sheets><sheet name="{_xml_escape(sheet_name)}" sheetId="1" r:id="rId1"/></sheets></workbook>'
    )

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("xl/workbook.xml", wb_xml)
        z.writestr("xl/_rels/workbook.xml.rels", WB_RELS)
        z.writestr("xl/sharedStrings.xml", sst_xml)
        z.writestr("xl/worksheets/sheet1.xml", sheet_xml)


def main():
    out = os.path.join(os.path.dirname(__file__), "sample_data")
    os.makedirs(out, exist_ok=True)

    left = [
        ["商品ID", "商品名", "価格", "分類", "備考"],
        ["A001", "商品A", 1000, "食品", "在庫あり"],
        ["A002", "商品B", 1200, "食品", ""],
        ["A003", "商品C", 800, "日用品", "廃番予定"],
        ["A004", "商品D", 1500, "家電", ""],
        ["A005", "商品E", 2000, "家電", "人気"],
        ["A006", "商品F", 500, "食品", "A001"],
    ]

    right = [
        ["商品ID", "商品名", "価格", "分類", "備考"],
        ["A001", "商品A", 1000, "食品", "在庫あり"],
        ["A002", "商品B", 1300, "食品", "値上げ"],   # 価格・備考が変更
        ["A004", "商品D", 1500, "家電", ""],          # A003 が削除された
        ["A005", "商品E改", 2000, "家電", "人気"],     # 商品名変更
        ["A007", "商品G", 900, "日用品", "新商品"],    # 追加
        ["A006", "商品F", 500, "食品", "A001"],
    ]

    # 3ファイル比較用（C）。left を別の担当者が編集した想定。
    third = [
        ["商品ID", "商品名", "価格", "分類", "備考"],
        ["A001", "商品A", 1000, "食品", "在庫あり"],
        ["A002", "商品B", 1200, "食品", "在庫僅少"],   # 備考のみ変更（B とも違う）
        ["A003", "商品C", 800, "日用品", "廃番予定"],
        ["A004", "商品D", 1600, "家電", ""],           # 価格変更（B は据え置き）
        ["A005", "商品E改", 2000, "家電", "人気"],      # B と同じ変更
        ["A007", "商品G", 900, "日用品", "新商品"],     # B と同じ追加行
        ["A006", "商品F", 500, "食品", "A001"],
    ]

    write_xlsx(os.path.join(out, "left.xlsx"), left)
    write_xlsx(os.path.join(out, "right.xlsx"), right)
    write_xlsx(os.path.join(out, "third.xlsx"), third)
    print("生成しました:", out)


if __name__ == "__main__":
    main()
