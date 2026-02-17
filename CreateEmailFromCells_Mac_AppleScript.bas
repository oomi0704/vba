' ============================================
' セルのデータを取得してメールを作成（Mac版 - AppleScript使用）
' Macのメールアプリを直接操作
' ============================================
Sub CreateEmailFromCells_AppleScript()
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
    
    ' AppleScriptでメールアプリを起動
    #If Mac Then
        Dim script As String
        ' 本文内の改行と引用符をエスケープ
        Dim escapedBody As String
        escapedBody = Replace(bodyText, vbCrLf, "\n")
        escapedBody = Replace(escapedBody, """", "\""")
        
        Dim escapedSubject As String
        escapedSubject = Replace(subject, """", "\""")
        
        ' メールアプリで新規メールを作成
        script = "tell application ""Mail""" & vbCrLf & _
                 "    activate" & vbCrLf & _
                 "    set newMessage to make new outgoing message with properties {" & _
                 "subject: """ & escapedSubject & """, " & _
                 "content: """ & escapedBody & """, " & _
                 "visible: true}" & vbCrLf & _
                 "    tell newMessage" & vbCrLf & _
                 "        make new to recipient at end of to recipients with properties {address: """ & recipient & """}" & vbCrLf & _
                 "    end tell" & vbCrLf & _
                 "end tell"
        
        MacScript script
        MsgBox "メールアプリでメールを作成しました。内容を確認して送信してください。", vbInformation, "完了"
    #Else
        MsgBox "このコードはMac環境でのみ動作します。", vbExclamation, "環境エラー"
    #End If
    
    Exit Sub
    
ErrorHandler:
    Dim errorMsg As String
    errorMsg = "エラーが発生しました:" & vbCrLf & vbCrLf
    errorMsg = errorMsg & "エラー番号: " & Err.Number & vbCrLf
    errorMsg = errorMsg & "エラー内容: " & Err.Description & vbCrLf & vbCrLf
    
    If Err.Number = 429 Then
        errorMsg = errorMsg & "メールアプリが起動できません。" & vbCrLf
        errorMsg = errorMsg & "Macのメールアプリがインストールされているか確認してください。"
    ElseIf Err.Number = 13 Then
        errorMsg = errorMsg & "データ型のエラーです。" & vbCrLf
        errorMsg = errorMsg & "セルの値が正しい形式か確認してください。"
    Else
        errorMsg = errorMsg & "予期しないエラーが発生しました。"
    End If
    
    MsgBox errorMsg, vbCritical, "エラー"
End Sub
