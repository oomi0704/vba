' ============================================
' セルのデータを取得してメールを作成（シンプル版）
' Mac環境でエラー5が発生する場合の代替方法
' ============================================
Sub CreateEmailFromCells_Simple()
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
    
    ' メール本文をクリップボードにコピー
    Dim dataObj As Object
    Set dataObj = CreateObject("MSForms.DataObject")
    dataObj.SetText bodyText
    dataObj.PutInClipboard
    
    ' mailto:リンクを作成（本文は短く）
    Dim mailtoLink As String
    mailtoLink = "mailto:" & recipient & "?subject=" & EncodeURL_Simple(subject)
    
    ' デフォルトのメールアプリで開く
    #If Mac Then
        ' Macの場合: より簡単な方法
        On Error Resume Next
        ' シェルコマンドを使用
        Shell "open """ & mailtoLink & """", vbHide
        If Err.Number <> 0 Then
            ' フォールバック: AppleScriptを1行で実行
            Dim simpleScript As String
            simpleScript = "do shell script ""open '" & Replace(mailtoLink, "'", "'\''") & "'"""
            MacScript simpleScript
        End If
        On Error GoTo ErrorHandler
    #Else
        ' Windowsの場合: ShellExecuteを使用
        Shell "cmd /c start """ & mailtoLink & """", vbHide
    #End If
    
    ' メッセージを表示
    Dim msg As String
    msg = "メールアプリを開きました。" & vbCrLf & vbCrLf
    msg = msg & "メール本文はクリップボードにコピーされています。" & vbCrLf
    msg = msg & "メールアプリで本文欄に貼り付けてください。"
    MsgBox msg, vbInformation, "完了"
    
    Exit Sub
    
ErrorHandler:
    Dim errorMsg As String
    errorMsg = "エラーが発生しました:" & vbCrLf & vbCrLf
    errorMsg = errorMsg & "エラー番号: " & Err.Number & vbCrLf
    errorMsg = errorMsg & "エラー内容: " & Err.Description & vbCrLf & vbCrLf
    
    Select Case Err.Number
        Case 5
            errorMsg = errorMsg & "プロシージャの呼び出しエラーです。" & vbCrLf & vbCrLf
            errorMsg = errorMsg & "【対処方法】" & vbCrLf
            errorMsg = errorMsg & "1. システム環境設定 → セキュリティとプライバシー → プライバシー" & vbCrLf
            errorMsg = errorMsg & "   「自動操作」にExcelを追加" & vbCrLf
            errorMsg = errorMsg & "2. Excelを再起動" & vbCrLf
            errorMsg = errorMsg & "3. 再度実行"
        Case 13
            errorMsg = errorMsg & "データ型のエラーです。" & vbCrLf
            errorMsg = errorMsg & "セルの値が正しい形式か確認してください。"
        Case Else
            errorMsg = errorMsg & "予期しないエラーが発生しました。"
    End Select
    
    MsgBox errorMsg, vbCritical, "エラー"
End Sub

' ============================================
' URLエンコード関数（シンプル版）
' ============================================
Function EncodeURL_Simple(text As String) As String
    Dim result As String
    Dim i As Integer
    Dim char As String
    
    result = ""
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        Select Case char
            Case " "
                result = result & "%20"
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
    
    EncodeURL_Simple = result
End Function
