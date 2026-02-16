# Excel VBA ガイド

## VBAファイルの使い方

### 1. ExcelでVBAを開く方法

1. Excelを開く
2. `Alt + F11` を押す（Mac: `Fn + Option + F11`）
3. 左側の「プロジェクト」ウィンドウで、ブックを右クリック
4. 「挿入」→「標準モジュール」を選択
5. 作成されたモジュールに、`sample_vba.bas` のコードをコピー＆ペースト

### 2. マクロの実行方法

1. `Alt + F8` を押す（Mac: `Fn + Option + F8`）
2. 実行したいマクロを選択
3. 「実行」をクリック

または、VBAエディタでカーソルをマクロ内に置いて `F5` を押す

### 3. マクロをボタンに割り当てる方法

1. 「開発」タブを表示（設定→リボンのユーザー設定→開発にチェック）
2. 「挿入」→「ボタン（フォームコントロール）」を選択
3. シート上でボタンを描画
4. マクロの割り当てダイアログで、実行したいマクロを選択

## サンプルコードの説明

### BasicCellOperation
- セルへの値の書き込み・読み込み
- 数式の設定
- 書式設定（太字、色）

### LoopExample
- For文を使ったループ処理
- 最後の行を自動検出してループ

### ConditionalExample
- If文を使った条件分岐
- セルの値に応じて判定

### DataSummary
- データの合計を計算
- 最後の行に合計を追加

### SheetOperation
- シートの作成・選択・コピー・削除

### FileOperation
- 新しいブックを作成して保存

### SendEmail
- Outlookと連携してメール送信

### ErrorHandlingExample
- エラーハンドリングの基本

### ArrayExample
- 配列の使い方
- セル範囲と配列の相互変換

### DateTimeExample
- 日付・時刻の取得と計算

### CreateChart
- グラフの作成

## よく使うVBAコード

### 最後の行を取得
```vba
Dim lastRow As Long
lastRow = Cells(Rows.Count, 1).End(xlUp).Row
```

### 最後の列を取得
```vba
Dim lastCol As Long
lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
```

### セル範囲を選択
```vba
Range("A1:C10").Select
```

### アクティブセルの行・列を取得
```vba
Dim currentRow As Long
Dim currentCol As Long
currentRow = ActiveCell.Row
currentCol = ActiveCell.Column
```

### メッセージボックスを表示
```vba
MsgBox "メッセージ"
```

### 入力ボックスで値を取得
```vba
Dim userInput As String
userInput = InputBox("値を入力してください")
```

## トラブルシューティング

### マクロが実行されない
- 「開発」タブで「マクロのセキュリティ」を確認
- 「すべてのマクロを有効にする」に設定（セキュリティに注意）

### エラーが発生する
- `Option Explicit` がある場合、変数は必ず宣言する
- エラーハンドリング（`On Error GoTo`）を追加

### Outlook連携が動かない
- Outlookがインストールされているか確認
- 「Microsoft Outlook xx.x Object Library」を参照設定に追加

## 参考リソース

- [Microsoft公式: Excel VBA リファレンス](https://learn.microsoft.com/ja-jp/office/vba/api/overview/excel)
- VBAエディタのヘルプ（F1キー）
