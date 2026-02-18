# Mac/Windows対応の比較

## 元のコード（Windows専用）

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

### このコードがMacで動作しない理由

1. **`New Outlook.Application`**
   - WindowsのCOMオブジェクトを使用
   - MacではCOMオブジェクトが使用できない

2. **参照設定が必要**
   - `Outlook.MailItem` という型を使うには参照設定が必要
   - Mac版Excelでは参照設定の動作が異なる

3. **ファイルパス**
   - `C:\Users\...` はWindows形式
   - Macでは `/Users/...` 形式

---

## Mac対応版の特徴

### 1. 参照設定不要

```vba
' Windows専用（参照設定必要）
Dim myOL As New Outlook.Application
Dim myOLMI As Outlook.MailItem

' Mac/Windows対応（参照設定不要）
Dim myOL As Object
Dim myOLMI As Object
Set myOL = CreateObject("Outlook.Application")
```

### 2. 環境自動判定

```vba
#If Mac Then
    ' Mac用のコード
    Call CreateEmailMac(...)
#Else
    ' Windows用のコード
    Call CreateEmailWindows(...)
#End If
```

### 3. Macではmailto:リンクを使用

**Mac環境：**
- mailto:リンクでメールアプリを開く
- デフォルトのメールアプリ（Mail.appなど）が起動
- CC、件名、本文を設定可能
- BCCと添付ファイルは制限あり

**Windows環境：**
- Outlook COMオブジェクトを使用
- 元のコードと同じ動作
- すべての機能が使用可能

---

## 機能比較

| 機能 | Windows版 | Mac版 |
|------|-----------|-------|
| 宛先（To） | ✅ | ✅ |
| CC | ✅ | ✅ |
| BCC | ✅ | ⚠️ 手動設定が必要 |
| 件名 | ✅ | ✅ |
| 本文 | ✅ | ✅（長い場合はクリップボード） |
| 添付ファイル | ✅ | ⚠️ 手動追加が必要 |

---

## Mac版の制限事項

### 1. BCCが設定できない

**理由：**
- mailto:リンクではBCCを設定できない

**対処方法：**
- メールアプリで手動でBCCを設定
- または、メール本文に「BCC: xxx@example.com」と記載して注意喚起

### 2. 添付ファイルが自動追加できない

**理由：**
- mailto:リンクでは添付ファイルを指定できない

**対処方法：**
- メールアプリで手動で添付ファイルを追加
- または、ファイルパスをメール本文に記載

### 3. 長い本文が設定できない場合がある

**理由：**
- mailto:リンクのURL長さに制限がある

**対処方法：**
- 本文をクリップボードにコピー
- メールアプリで貼り付け

---

## 使い方

### Excelシートの設定

「メール設定」シート：

| A列 | B列 |
|-----|-----|
| 宛先 | recipient@example.com |
| CC | cc@example.com |
| BCC | bcc@example.com |
| 件名 | 月次報告 |
| 本文 | お世話になっております。... |
| 添付ファイル | /Users/username/Desktop/sample.xlsx |

**注意：**
- Mac環境では、添付ファイルのパスは `/Users/...` 形式
- Windows環境では、`C:\Users\...` 形式

---

## 推奨事項

### Mac環境で完全な機能が必要な場合

1. **Mac版Outlookを使用**
   - Mac版Outlookがインストールされている場合、COMオブジェクトが使える可能性がある
   - ただし、動作が保証されていない

2. **AppleScriptを使用**
   - Macのメールアプリを直接操作
   - より高度な制御が可能
   - ただし、コードが複雑になる

3. **現在のMac対応版を使用**
   - 最もシンプルで確実
   - BCCと添付ファイルは手動設定が必要

---

## まとめ

- **元のコード**: Windows専用、参照設定必要
- **Mac対応版**: Mac/Windows両対応、参照設定不要
- **Mac版の制限**: BCCと添付ファイルは手動設定が必要

Mac環境では、`Main1_Mac.bas` のコードを使用してください。
