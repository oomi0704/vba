# VBA Project

## 概要

VBAの勉強をするためのプロジェクトです。PythonでExcelを操作するサンプルコードも含まれています。

## セットアップ

### Python環境のセットアップ

```bash
# 仮想環境を作成
python3 -m venv venv

# 仮想環境をアクティベート
source venv/bin/activate  # Mac/Linux
# または
# venv\Scripts\activate  # Windows

# 依存関係をインストール
pip install --trusted-host pypi.org --trusted-host pypi.python.org --trusted-host files.pythonhosted.org -r requirements.txt
```

## 使用方法

### Excel操作スクリプトの実行

```bash
# 仮想環境をアクティベート
source venv/bin/activate

# スクリプトを実行
python excel_operation.py
```

実行すると `output.xlsx` が作成されます。

詳細は `EXCEL_LINKAGE.md` を参照してください。
