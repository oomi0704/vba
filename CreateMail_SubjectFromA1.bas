' ============================================
' Sheet1のA1セルから件名を取得してメールを作成
' Mac/Windows対応版
' ============================================
Sub CreateMail_SubjectFromA1()
    On Error GoTo ErrorHandler
    
    Dim subjectText As String
    Dim bodyText As String
    Dim recipient As String
    
    ' Sheet1のA1セルの内容を取得
    subjectText = Trim(ThisWorkbook.Sheets("Sheet1").Range("A1").Value)
    
    ' 件名が空の場合はエラー
    If subjectText = "" Then
        MsgBox "エラー: Sheet1のA1セルが空です。", vbExclamation, "入力エラー"
        Exit Sub
    End If
    
    ' 本文と宛先を設定（必要に応じて変更）
    bodyText = "本文をここに入力できます"
    recipient = ""  ' 宛先があれば入力
    
    ' Mac/Windowsで処理を分岐
    #If Mac Then
        ' Macの場合: mailto:リンクを使用
        Call CreateEmailMac_Simple(recipient, subjectText, bodyText)
    #Else
        ' Windowsの場合: Outlook COMオブジェクトを使用
        Call CreateEmailWindows_Simple(recipient, subjectText, bodyText)
    #End If
    
    Exit Sub
    
ErrorHandler:
    Dim errorMsg As String
    errorMsg = "エラーが発生しました:" & vbCrLf & vbCrLf
    errorMsg = errorMsg & "エラー番号: " & Err.Number & vbCrLf
    errorMsg = errorMsg & "エラー内容: " & Err.Description & vbCrLf & vbCrLf
    
    Select Case Err.Number
        Case 9
            errorMsg = errorMsg & "Sheet1が見つかりません。" & vbCrLf
            errorMsg = errorMsg & "シート名を確認してください。"
        Case 429
            errorMsg = errorMsg & "Outlookが起動できません。" & vbCrLf & vbCrLf
            errorMsg = errorMsg & "【Mac環境の場合】" & vbCrLf
            errorMsg = errorMsg & "このコードはMac対応版に自動的に切り替わります。" & vbCrLf & vbCrLf
            errorMsg = errorMsg & "【Windows環境の場合】" & vbCrLf
            errorMsg = errorMsg & "Outlookがインストールされているか確認してください。"
        Case Else
            errorMsg = errorMsg & "予期しないエラーが発生しました。"
    End Select
    
    MsgBox errorMsg, vbCritical, "エラー"
End Sub

' ============================================
' Mac用メール作成（シンプル版）
' ============================================
Sub CreateEmailMac_Simple(recipient As String, subject As String, body As String)
    
    ' メール本文をクリップボードにコピー
    Dim dataObj As Object
    On Error Resume Next
    Set dataObj = CreateObject("MSForms.DataObject")
    If Err.Number = 0 Then
        dataObj.SetText body
        dataObj.PutInClipboard
    End If
    On Error GoTo 0
    
    ' mailto:リンクを作成
    Dim mailtoLink As String
    If recipient <> "" Then
        mailtoLink = "mailto:" & recipient
    Else
        mailtoLink = "mailto:"
    End If
    
    ' 件名を追加
    If subject <> "" Then
        mailtoLink = mailtoLink & "?subject=" & EncodeURL_Simple(subject)
    End If
    
    ' 本文を追加（短い場合のみ）
    If Len(body) < 500 And body <> "" Then
        If InStr(mailtoLink, "?") > 0 Then
            mailtoLink = mailtoLink & "&body=" & EncodeURL_Simple(body)
        Else
            mailtoLink = mailtoLink & "?body=" & EncodeURL_Simple(body)
        End If
    End If
    
    ' メールアプリを開く
    On Error Resume Next
    Shell "open """ & mailtoLink & """", vbHide
    If Err.Number <> 0 Then
        ' フォールバック: AppleScriptを使用
        Dim script As String
        script = "do shell script ""open '" & Replace(mailtoLink, "'", "'\''") & "'"""
        MacScript script
    End If
    On Error GoTo 0
    
    ' メッセージを表示
    Dim msg As String
    msg = "メールアプリを開きました。" & vbCrLf & vbCrLf
    
    If Len(body) >= 500 Then
        msg = msg & "メール本文はクリップボードにコピーされています。" & vbCrLf
        msg = msg & "本文欄に貼り付けてください。"
    End If
    
    MsgBox msg, vbInformation, "完了"
End Sub

' ============================================
' Windows用メール作成（元のコードを改善）
' ============================================
Sub CreateEmailWindows_Simple(recipient As String, subject As String, body As String)
    
    Dim OutApp As Object
    Dim OutMail As Object
    
    ' Outlook を起動（参照設定不要の方法）
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)  ' 0 = olMailItem
    
    With OutMail
        If recipient <> "" Then .To = recipient
        .Subject = subject
        .Body = body
        .Display  ' 送信前に確認
    End With
    
    ' クリーンアップ
    Set OutMail = Nothing
    Set OutApp = Nothing
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
    
    EncodeURL_Simple = result
End Function
