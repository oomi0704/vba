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
