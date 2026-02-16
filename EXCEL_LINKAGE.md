# Excel と Python の連携方法

## 方法1: openpyxl（ファイル操作）

### 特徴
- ✅ Excel がインストールされていなくても動作
- ✅ ファイルを直接読み書き
- ✅ 軽量で高速
- ❌ Excel アプリケーションの機能（マクロ実行など）は使えない

### 使い方

```python
from openpyxl import Workbook, load_workbook

# 新規作成
wb = Workbook()
ws = wb.active
ws["A1"] = "Hello"
wb.save("output.xlsx")

# 既存ファイルを開く
wb = load_workbook("existing.xlsx")
ws = wb.active
value = ws["A1"].value
wb.close()
```

### インストール
```bash
pip install openpyxl
```

---

## 方法2: xlwings（Excel アプリケーション操作）

### 特徴
- ✅ Excel アプリケーションを直接操作
- ✅ マクロを実行できる
- ✅ 数式、グラフ、VBA との連携が可能
- ✅ Windows/Mac 両方対応
- ❌ Excel がインストールされている必要がある

### 使い方

```python
import xlwings as xw

# Excel を起動
app = xw.App(visible=True)

# ブックを開く
wb = app.books.open("file.xlsx")
ws = wb.sheets[0]

# セル操作
ws.range("A1").value = "Hello"

# マクロ実行
app.api.Run("Module1.MyMacro")

# 保存して閉じる
wb.save()
wb.close()
app.quit()
```

### インストール
```bash
pip install xlwings
```

---

## 方法3: pandas（データ分析向け）

### 特徴
- ✅ データ分析・集計に最適
- ✅ CSV、Excel、JSON など様々な形式に対応
- ✅ データ処理が簡単

### 使い方

```python
import pandas as pd

# Excel を読み込む
df = pd.read_excel("data.xlsx", sheet_name="Sheet1")

# データ処理
df["合計"] = df["数量"] * df["単価"]
summary = df.groupby("商品名").sum()

# Excel に書き出す
summary.to_excel("output.xlsx", sheet_name="集計")
```

### インストール
```bash
pip install pandas openpyxl
```

---

## 使い分けの目安

| 用途 | 推奨ライブラリ |
|------|----------------|
| ファイルの読み書きだけ | **openpyxl** |
| Excel アプリケーションを操作したい | **xlwings** |
| データ分析・集計 | **pandas** |
| マクロを実行したい | **xlwings** |
| Excel が入っていない環境 | **openpyxl** または **pandas** |

---

## 実行例

```bash
# 依存関係をインストール
pip install -r requirements.txt

# スクリプトを実行
python excel_operation.py
```
