"""
Excel を操作する簡単な Python スクリプト
openpyxl を使用（Excel がインストールされていなくても動作）
"""

from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

# 出力ファイル名
OUTPUT_FILE = "output.xlsx"


def create_and_write():
    """新しい Excel ブックを作成して書き込む"""
    wb = Workbook()
    ws = wb.active
    ws.title = "サンプル"

    # セルに値を書き込む
    ws["A1"] = "商品名"
    ws["B1"] = "数量"
    ws["C1"] = "単価"
    ws["D1"] = "合計"

    # 見出しを太字に
    for col in ["A1", "B1", "C1", "D1"]:
        ws[col].font = Font(bold=True)

    # データを書き込む
    data = [
        ["りんご", 5, 120],
        ["みかん", 10, 80],
        ["バナナ", 3, 150],
    ]
    for i, row in enumerate(data, start=2):
        ws.cell(row=i, column=1, value=row[0])
        ws.cell(row=i, column=2, value=row[1])
        ws.cell(row=i, column=3, value=row[2])
        ws.cell(row=i, column=4, value=row[1] * row[2])  # 合計

    # 列幅を調整
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 8
    ws.column_dimensions["D"].width = 10

    wb.save(OUTPUT_FILE)
    print(f"✅ {OUTPUT_FILE} を作成しました。")


def read_excel(file_path):
    """既存の Excel ファイルを読み込んで内容を表示"""
    wb = load_workbook(file_path, read_only=False)
    ws = wb.active
    print(f"\n📖 シート名: {ws.title}\n")

    for row in ws.iter_rows(min_row=1, values_only=True):
        print(row)

    wb.close()


if __name__ == "__main__":
    # 1. 新規作成して保存
    create_and_write()

    # 2. 保存したファイルを読み込んで表示
    read_excel(OUTPUT_FILE)
