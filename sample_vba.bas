' ============================================
' Excel VBA サンプルコード
' ============================================

Option Explicit

' ============================================
' 基本的なセル操作
' ============================================
Sub BasicCellOperation()
    ' セルに値を書き込む
    Range("A1").Value = "Hello"
    Range("B1").Value = "World"
    
    ' セルの値を読み込む
    Dim value As String
    value = Range("A1").Value
    MsgBox "A1の値: " & value
    
    ' セルに数式を設定
    Range("C1").Formula = "=A1&B1"
    
    ' セルの書式設定
    Range("A1").Font.Bold = True
    Range("A1").Font.Color = RGB(255, 0, 0)  ' 赤色
End Sub

' ============================================
' ループ処理
' ============================================
Sub LoopExample()
    Dim i As Integer
    
    ' 1から10までループ
    For i = 1 To 10
        Cells(i, 1).Value = i
        Cells(i, 2).Value = i * 2
    Next i
    
    ' 最後の行までループ
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim j As Long
    For j = 1 To lastRow
        ' 処理をここに記述
        Debug.Print Cells(j, 1).Value
    Next j
End Sub

' ============================================
' 条件分岐
' ============================================
Sub ConditionalExample()
    Dim score As Integer
    score = Range("A1").Value
    
    If score >= 80 Then
        Range("B1").Value = "合格"
        Range("B1").Font.Color = RGB(0, 255, 0)  ' 緑色
    ElseIf score >= 60 Then
        Range("B1").Value = "要努力"
        Range("B1").Font.Color = RGB(255, 165, 0)  ' オレンジ色
    Else
        Range("B1").Value = "不合格"
        Range("B1").Font.Color = RGB(255, 0, 0)  ' 赤色
    End If
End Sub

' ============================================
' データの集計
' ============================================
Sub DataSummary()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' 最後の行を取得
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    ' 合計を計算
    Dim total As Double
    Dim i As Long
    For i = 2 To lastRow  ' 1行目は見出しと仮定
        total = total + ws.Cells(i, 2).Value
    Next i
    
    ' 合計を表示
    ws.Cells(lastRow + 1, 1).Value = "合計"
    ws.Cells(lastRow + 1, 2).Value = total
    ws.Cells(lastRow + 1, 2).Font.Bold = True
End Sub

' ============================================
' シート操作
' ============================================
Sub SheetOperation()
    ' 新しいシートを作成
    Dim newSheet As Worksheet
    Set newSheet = Worksheets.Add
    newSheet.Name = "新しいシート"
    
    ' シートを選択
    Worksheets("Sheet1").Select
    
    ' シートをコピー
    Worksheets("Sheet1").Copy After:=Worksheets(Worksheets.Count)
    
    ' シートを削除（確認付き）
    Application.DisplayAlerts = False
    Worksheets("新しいシート").Delete
    Application.DisplayAlerts = True
End Sub

' ============================================
' ファイル操作
' ============================================
Sub FileOperation()
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\output.xlsx"
    
    ' 新しいブックを作成して保存
    Dim newBook As Workbook
    Set newBook = Workbooks.Add
    newBook.SaveAs Filename:=filePath
    newBook.Close
    
    MsgBox "ファイルを保存しました: " & filePath
End Sub

' ============================================
' メール送信（Outlook連携）
' ============================================
Sub SendEmail()
    Dim olApp As Object
    Dim olMail As Object
    
    ' Outlook を起動
    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)  ' 0 = メールアイテム
    
    ' メールの設定
    olMail.To = "recipient@example.com"
    olMail.Subject = "件名"
    olMail.Body = "本文です。" & vbCrLf & "Excel VBAから送信しています。"
    
    ' 送信（確認なしで送信する場合）
    ' olMail.Send
    
    ' 確認してから送信する場合
    olMail.Display
End Sub

' ============================================
' セルのデータを取得してメールを作成（Mac/Windows対応版）
' mailto:リンクを使用（どの環境でも動作）
' エラーハンドリングと空セルチェック付き
' ============================================
Sub CreateEmailFromCells()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' セルからデータを取得
    Dim recipient As String
    Dim subject As String
    Dim name As String
    Dim amount As Variant
    Dim dateValue As Variant
    
    ' セルの位置を指定（必要に応じて変更）
    recipient = Trim(ws.Range("B2").Value)  ' 宛先
    subject = Trim(ws.Range("B3").Value)     ' 件名
    name = Trim(ws.Range("B4").Value)        ' 名前
    amount = ws.Range("B5").Value            ' 金額
    dateValue = ws.Range("B6").Value        ' 日付
    
    ' 必須項目のチェック
    If recipient = "" Then
        MsgBox "エラー: B2セル（宛先）が空です。", vbExclamation, "入力エラー"
        Exit Sub
    End If
    
    If subject = "" Then
        MsgBox "エラー: B3セル（件名）が空です。", vbExclamation, "入力エラー"
        Exit Sub
    End If
    
    If name = "" Then
        MsgBox "エラー: B4セル（名前）が空です。", vbExclamation, "入力エラー"
        Exit Sub
    End If
    
    ' メール本文を作成
    Dim bodyText As String
    bodyText = name & " 様" & vbCrLf & vbCrLf
    bodyText = bodyText & "お世話になっております。" & vbCrLf & vbCrLf
    bodyText = bodyText & "以下の内容をご確認ください。" & vbCrLf & vbCrLf
    
    ' 金額が入力されている場合のみ表示
    If IsNumeric(amount) And amount <> "" Then
        bodyText = bodyText & "金額: " & Format(amount, "#,##0") & "円" & vbCrLf
    End If
    
    ' 日付が入力されている場合のみ表示
    If IsDate(dateValue) And dateValue <> "" Then
        bodyText = bodyText & "日付: " & Format(dateValue, "yyyy年mm月dd日") & vbCrLf
    End If
    
    bodyText = bodyText & vbCrLf & "よろしくお願いいたします。"
    
    ' mailto:リンクを作成
    Dim mailtoLink As String
    mailtoLink = "mailto:" & recipient & "?subject=" & EncodeURL(subject) & "&body=" & EncodeURL(bodyText)
    
    ' デフォルトのメールアプリで開く
    #If Mac Then
        ' Macの場合: AppleScriptを使用（1行形式）
        Dim script As String
        ' 引用符をエスケープ
        Dim escapedLink As String
        escapedLink = Replace(mailtoLink, """", "\""")
        script = "tell application ""System Events"" to open location """ & escapedLink & """"
        MacScript script
    #Else
        ' Windowsの場合: ShellExecuteを使用
        Shell "cmd /c start """ & mailtoLink & """", vbHide
    #End If
    
    MsgBox "メールアプリを開きました。内容を確認して送信してください。", vbInformation, "完了"
    
    Exit Sub
    
ErrorHandler:
    Dim errorMsg As String
    errorMsg = "エラーが発生しました:" & vbCrLf & vbCrLf
    errorMsg = errorMsg & "エラー番号: " & Err.Number & vbCrLf
    errorMsg = errorMsg & "エラー内容: " & Err.Description & vbCrLf & vbCrLf
    
    Select Case Err.Number
        Case 5
            errorMsg = errorMsg & "プロシージャの呼び出しエラーです。" & vbCrLf & vbCrLf
            errorMsg = errorMsg & "【Mac環境の場合】" & vbCrLf
            errorMsg = errorMsg & "システム環境設定 → セキュリティとプライバシー → プライバシー" & vbCrLf
            errorMsg = errorMsg & "「自動操作」にExcelを追加してください。" & vbCrLf & vbCrLf
            errorMsg = errorMsg & "または、mailto:リンクが長すぎる可能性があります。" & vbCrLf
            errorMsg = errorMsg & "メール本文を短くするか、別の方法をお試しください。"
        Case 13
            errorMsg = errorMsg & "データ型のエラーです。" & vbCrLf
            errorMsg = errorMsg & "セルの値が正しい形式か確認してください。"
        Case Else
            errorMsg = errorMsg & "予期しないエラーが発生しました。"
    End Select
    
    MsgBox errorMsg, vbCritical, "エラー"
End Sub

' ============================================
' URLエンコード関数（mailto:リンク用）
' ============================================
Function EncodeURL(text As String) As String
    Dim result As String
    Dim i As Integer
    Dim char As String
    
    result = ""
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        Select Case char
            Case " "
                result = result & "%20"
            Case vbCrLf
                result = result & "%0D%0A"
            Case "&"
                result = result & "%26"
            Case "="
                result = result & "%3D"
            Case "?"
                result = result & "%3F"
            Case "#"
                result = result & "%23"
            Case "%"
                result = result & "%25"
            Case "+"
                result = result & "%2B"
            Case Else
                result = result & char
        End Select
    Next i
    
    EncodeURL = result
End Function

' ============================================
' テーブル形式のデータを取得してメールを作成
' ============================================
Sub CreateEmailFromTable()
    Dim olApp As Object
    Dim olMail As Object
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Outlook を起動
    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)
    
    ' メールの基本情報をセルから取得
    olMail.To = ws.Range("B1").Value      ' 宛先（B1セル）
    olMail.Subject = ws.Range("B2").Value ' 件名（B2セル）
    
    ' データテーブルの範囲を取得（A4から始まると仮定）
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' メール本文を作成
    Dim bodyText As String
    bodyText = "お世話になっております。" & vbCrLf & vbCrLf
    bodyText = bodyText & "以下のデータをご確認ください。" & vbCrLf & vbCrLf
    
    ' テーブルの見出しを取得（4行目を想定）
    Dim headerRow As Integer
    headerRow = 4
    
    ' 見出しを取得
    Dim headers As String
    Dim col As Integer
    headers = ""
    For col = 1 To 4  ' A列からD列まで
        If col > 1 Then headers = headers & " | "
        headers = headers & ws.Cells(headerRow, col).Value
    Next col
    bodyText = bodyText & headers & vbCrLf
    bodyText = bodyText & String(Len(headers), "-") & vbCrLf
    
    ' データ行を取得
    Dim i As Long
    For i = headerRow + 1 To lastRow
        Dim rowData As String
        rowData = ""
        For col = 1 To 4
            If col > 1 Then rowData = rowData & " | "
            rowData = rowData & ws.Cells(i, col).Value
        Next col
        bodyText = bodyText & rowData & vbCrLf
    Next i
    
    bodyText = bodyText & vbCrLf & "よろしくお願いいたします。"
    
    ' メール本文を設定
    olMail.Body = bodyText
    
    ' メールを表示
    olMail.Display
End Sub

' ============================================
' 各行ごとに個別のメールを作成
' ============================================
Sub CreateMultipleEmails()
    Dim olApp As Object
    Dim olMail As Object
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Outlook を起動
    Set olApp = CreateObject("Outlook.Application")
    
    ' データの開始行（2行目からデータがあると仮定、1行目は見出し）
    Dim startRow As Long
    Dim lastRow As Long
    startRow = 2
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim i As Long
    For i = startRow To lastRow
        ' 各行からデータを取得
        Dim recipient As String
        Dim name As String
        Dim amount As Double
        Dim dateValue As Date
        
        recipient = ws.Cells(i, 1).Value  ' A列: 宛先メールアドレス
        name = ws.Cells(i, 2).Value       ' B列: 名前
        amount = ws.Cells(i, 3).Value      ' C列: 金額
        dateValue = ws.Cells(i, 4).Value   ' D列: 日付
        
        ' 空の行はスキップ
        If recipient = "" Then GoTo NextRow
        
        ' メールを作成
        Set olMail = olApp.CreateItem(0)
        
        ' メール本文を作成
        Dim bodyText As String
        bodyText = name & " 様" & vbCrLf & vbCrLf
        bodyText = bodyText & "お世話になっております。" & vbCrLf & vbCrLf
        bodyText = bodyText & "以下の内容をご確認ください。" & vbCrLf & vbCrLf
        bodyText = bodyText & "金額: " & Format(amount, "#,##0") & "円" & vbCrLf
        bodyText = bodyText & "日付: " & Format(dateValue, "yyyy年mm月dd日") & vbCrLf & vbCrLf
        bodyText = bodyText & "よろしくお願いいたします。"
        
        ' メールの設定
        olMail.To = recipient
        olMail.Subject = name & "様へのご連絡"
        olMail.Body = bodyText
        
        ' メールを表示
        olMail.Display
        
NextRow:
    Next i
    
    MsgBox lastRow - startRow + 1 & "件のメールを作成しました。"
End Sub

' ============================================
' Excelの範囲をHTML形式でメールに貼り付け
' ============================================
Sub CreateEmailWithExcelTable()
    Dim olApp As Object
    Dim olMail As Object
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Outlook を起動
    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)
    
    ' メールの基本設定
    olMail.To = ws.Range("B1").Value      ' 宛先
    olMail.Subject = ws.Range("B2").Value ' 件名
    
    ' Excelの範囲を指定
    Dim dataRange As Range
    Set dataRange = ws.Range("A4:D10")  ' 必要に応じて範囲を変更
    
    ' 範囲をコピー
    dataRange.Copy
    
    ' メール本文を作成
    Dim bodyText As String
    bodyText = "お世話になっております。" & vbCrLf & vbCrLf
    bodyText = bodyText & "以下のデータをご確認ください。" & vbCrLf & vbCrLf
    olMail.Body = bodyText
    
    ' HTML形式でメール本文に貼り付け
    olMail.HTMLBody = bodyText & "<br><br>"
    olMail.GetInspector
    
    ' クリップボードから貼り付け（手動操作をシミュレート）
    ' 注意: この方法はOutlookのバージョンによって動作が異なる場合があります
    Application.Wait Now + TimeValue("00:00:01")  ' 少し待機
    SendKeys "^v", True  ' Ctrl+V を送信
    
    ' より確実な方法: 範囲をHTMLに変換
    ' この場合は、範囲をHTMLテーブルとして作成する必要があります
    
    ' メールを表示
    olMail.Display
End Sub

' ============================================
' エラーハンドリング
' ============================================
Sub ErrorHandlingExample()
    On Error GoTo ErrorHandler
    
    Dim value As Double
    value = Range("A1").Value / Range("B1").Value
    
    Range("C1").Value = value
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description
    Range("C1").Value = "エラー"
End Sub

' ============================================
' 配列の使用
' ============================================
Sub ArrayExample()
    Dim data(1 To 5) As String
    Dim i As Integer
    
    ' 配列に値を設定
    data(1) = "りんご"
    data(2) = "みかん"
    data(3) = "バナナ"
    data(4) = "ぶどう"
    data(5) = "いちご"
    
    ' 配列の値をセルに書き込む
    For i = 1 To 5
        Cells(i, 1).Value = data(i)
    Next i
    
    ' セル範囲から配列に読み込む
    Dim cellData As Variant
    cellData = Range("A1:A5").Value
End Sub

' ============================================
' ユーザーフォームの表示
' ============================================
Sub ShowUserForm()
    ' UserForm1 というフォームがある場合
    ' UserForm1.Show
End Sub

' ============================================
' 日付・時刻の操作
' ============================================
Sub DateTimeExample()
    ' 現在の日時を取得
    Range("A1").Value = Now
    Range("A2").Value = Date
    Range("A3").Value = Time
    
    ' 日付の計算
    Range("B1").Value = DateAdd("d", 7, Date)  ' 7日後
    Range("B2").Value = DateAdd("m", 1, Date)  ' 1ヶ月後
    Range("B3").Value = DateAdd("yyyy", 1, Date)  ' 1年後
End Sub

' ============================================
' グラフの作成
' ============================================
Sub CreateChart()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' データ範囲を指定
    Dim chartRange As Range
    Set chartRange = ws.Range("A1:B5")
    
    ' グラフを作成
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=250, Top:=10, Width:=400, Height:=250)
    
    With chartObj.Chart
        .SetSourceData Source:=chartRange
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "サンプルグラフ"
    End With
End Sub
