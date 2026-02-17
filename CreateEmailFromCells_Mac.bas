' ============================================
' セルのデータを取得してメールを作成（Mac対応版）
' mailto:リンクを使用（どの環境でも動作）
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
        ' Macの場合: AppleScriptを使用
        Dim script As String
        script = "tell application ""System Events""" & vbCrLf & _
                 "    open location """ & mailtoLink & """" & vbCrLf & _
                 "end tell"
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
