# VBAコード解説: Outlookメール作成

## コード全体

```vba
Sub Main1() '参照設定あり
    Dim myOL As New Outlook.Application
    Dim myOLMI As Outlook.MailItem
    Dim settingSh As String
    
    settingSh = "メール設定"
    Set myOLMI = myOL.CreateItem(olMailItem)
    
    With myOLMI
        .To = ThisWorkbook.Sheets(settingSh).Range("B1")
        .CC = ThisWorkbook.Sheets(settingSh).Range("B2")
        .BCC = ThisWorkbook.Sheets(settingSh).Range("B3")
        .subject = ThisWorkbook.Sheets(settingSh).Range("B4")
        .Body = ThisWorkbook.Sheets(settingSh).Range("B5")
        .Attachments.Add "C:\Users\lenoco\Desktop\sample.xlsx"
        .Display
    End With
End Sub
```

---

## 行ごとの解説

### 1行目: `Sub Main1() '参照設定あり`

```vba
Sub Main1() '参照設定あり
```

- **`Sub Main1()`**: サブプロシージャ（関数）の定義。名前は `Main1`
- **`'参照設定あり`**: コメント。このコードは参照設定が必要であることを示す

**参照設定とは：**
- VBAエディタで「ツール」→「参照設定」
- 「Microsoft Outlook xx.x Object Library」にチェックを入れる必要がある
- これにより、`Outlook.Application` や `Outlook.MailItem` という型が使えるようになる

---

### 2行目: `Dim myOL As New Outlook.Application`

```vba
Dim myOL As New Outlook.Application
```

- **`Dim`**: 変数を宣言
- **`myOL`**: 変数名（Outlookアプリケーションオブジェクト）
- **`As New Outlook.Application`**: Outlookアプリケーションの新しいインスタンスを作成
  - `New` キーワードにより、宣言と同時にオブジェクトが作成される
  - Outlookアプリケーションを起動する

**注意：**
- Mac環境では動作しない（Windows専用）
- Outlookがインストールされている必要がある

---

### 3行目: `Dim myOLMI As Outlook.MailItem`

```vba
Dim myOLMI As Outlook.MailItem
```

- **`myOLMI`**: 変数名（メールアイテムオブジェクト）
- **`As Outlook.MailItem`**: Outlookのメールアイテム型を指定
- この時点ではまだ作成されていない（次の行で作成される）

---

### 4行目: `Dim settingSh As String`

```vba
Dim settingSh As String
```

- **`settingSh`**: 変数名（シート名を格納）
- **`As String`**: 文字列型

---

### 5行目: `settingSh = "メール設定"`

```vba
settingSh = "メール設定"
```

- シート名「メール設定」を変数に代入
- このシートからメールの設定を読み込む

---

### 6行目: `Set myOLMI = myOL.CreateItem(olMailItem)`

```vba
Set myOLMI = myOL.CreateItem(olMailItem)
```

- **`myOL.CreateItem(olMailItem)`**: 新しいメールアイテムを作成
  - `olMailItem` は定数で、値は `0`（メールアイテムを意味する）
- **`Set`**: オブジェクト変数に代入する際に必要
- これで新しいメールが作成される

**他のアイテムタイプ：**
- `olAppointmentItem` (1): 予定表アイテム
- `olContactItem` (2): 連絡先アイテム
- `olTaskItem` (3): タスクアイテム
- `olJournalItem` (4): ジャーナルアイテム
- `olNoteItem` (5): メモアイテム

---

### 8-15行目: `With myOLMI ... End With`

```vba
With myOLMI
    .To = ThisWorkbook.Sheets(settingSh).Range("B1")
    .CC = ThisWorkbook.Sheets(settingSh).Range("B2")
    .BCC = ThisWorkbook.Sheets(settingSh).Range("B3")
    .subject = ThisWorkbook.Sheets(settingSh).Range("B4")
    .Body = ThisWorkbook.Sheets(settingSh).Range("B5")
    .Attachments.Add "C:\Users\lenoco\Desktop\sample.xlsx"
    .Display
End With
```

**`With` ステートメント：**
- 同じオブジェクト（`myOLMI`）を繰り返し参照する際に便利
- `.To` は `myOLMI.To` の省略形

#### 各プロパティの説明

**`.To = ThisWorkbook.Sheets(settingSh).Range("B1")`**
- **`.To`**: 宛先（To）を設定
- **`ThisWorkbook`**: 現在のExcelブック
- **`.Sheets(settingSh)`**: 「メール設定」シートを参照
- **`.Range("B1")`**: B1セルの値を取得
- B1セルにメールアドレスが入っている想定

**`.CC = ThisWorkbook.Sheets(settingSh).Range("B2")`**
- **`.CC`**: CC（カーボンコピー）を設定
- B2セルの値をCCに設定

**`.BCC = ThisWorkbook.Sheets(settingSh).Range("B3")`**
- **`.BCC`**: BCC（ブラインドカーボンコピー）を設定
- B3セルの値をBCCに設定

**`.subject = ThisWorkbook.Sheets(settingSh).Range("B4")`**
- **`.subject`**: メールの件名を設定
- B4セルの値を件名に設定

**`.Body = ThisWorkbook.Sheets(settingSh).Range("B5")`**
- **`.Body`**: メール本文を設定
- B5セルの値を本文に設定

**`.Attachments.Add "C:\Users\lenoco\Desktop\sample.xlsx"`**
- **`.Attachments.Add`**: 添付ファイルを追加
- 指定したファイルパスのファイルを添付
- ファイルが存在しない場合はエラーになる

**`.Display`**
- メール作成ウィンドウを表示
- ユーザーが確認してから送信できる
- 自動送信する場合は `.Send` を使用

---

## Excelシートの設定例

「メール設定」シートに以下のように設定：

| A列 | B列 |
|-----|-----|
| 宛先 | recipient@example.com |
| CC | cc@example.com |
| BCC | bcc@example.com |
| 件名 | 月次報告 |
| 本文 | お世話になっております。... |

---

## このコードの特徴

### メリット
- ✅ シートから設定を読み込むので、コードを変更せずに設定を変更できる
- ✅ CC、BCC、添付ファイルにも対応
- ✅ 参照設定を使うので、コード補完が効く（IntelliSense）

### デメリット
- ❌ Windows専用（Macでは動作しない）
- ❌ Outlookがインストールされている必要がある
- ❌ 参照設定が必要
- ❌ ファイルパスが固定（`C:\Users\lenoco\Desktop\sample.xlsx`）

---

## 改善案

### 1. エラーハンドリングを追加

```vba
Sub Main1()
    On Error GoTo ErrorHandler
    
    Dim myOL As New Outlook.Application
    ' ... 既存のコード ...
    
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub
```

### 2. ファイルパスを動的にする

```vba
Dim filePath As String
filePath = ThisWorkbook.Path & "\sample.xlsx"
.Attachments.Add filePath
```

### 3. Mac対応版

```vba
#If Mac Then
    ' Mac用のコード（mailto:リンクなど）
#Else
    ' Windows用のコード（既存のコード）
#End If
```

---

## まとめ

このコードは：
1. Outlookアプリケーションを起動
2. 新しいメールを作成
3. Excelシートから設定を読み込む
4. 添付ファイルを追加
5. メール作成ウィンドウを表示

**使用するには：**
- Outlookがインストールされていること（Windows環境）
- 参照設定で「Microsoft Outlook Object Library」を有効化
- 「メール設定」シートに必要な情報を入力
