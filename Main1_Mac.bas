' ============================================
' Outlookメール作成（Mac対応版）
' 参照設定不要、Mac/Windows両対応
' ============================================
Sub Main1()
    On Error GoTo ErrorHandler
    
    Dim settingSh As String
    settingSh = "メール設定"
    
    ' シートから設定を読み込む
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(settingSh)
    
    Dim recipient As String
    Dim cc As String
    Dim bcc As String
    Dim subject As String
    Dim body As String
    Dim attachmentPath As String
    
    recipient = Trim(ws.Range("B1").Value)
    cc = Trim(ws.Range("B2").Value)
    bcc = Trim(ws.Range("B3").Value)
    subject = Trim(ws.Range("B4").Value)
    body = Trim(ws.Range("B5").Value)
    attachmentPath = Trim(ws.Range("B6").Value)  ' 添付ファイルのパス（オプション）
    
    ' 必須項目のチェック
    If recipient = "" Then
        MsgBox "エラー: B1セル（宛先）が空です。", vbExclamation, "入力エラー"
        Exit Sub
    End If
    
    If subject = "" Then
        MsgBox "エラー: B4セル（件名）が空です。", vbExclamation, "入力エラー"
        Exit Sub
    End If
    
    ' Mac/Windowsで処理を分岐
    #If Mac Then
        ' Macの場合: mailto:リンクを使用
        Call CreateEmailMac(recipient, cc, bcc, subject, body, attachmentPath)
    #Else
        ' Windowsの場合: Outlook COMオブジェクトを使用
        Call CreateEmailWindows(recipient, cc, bcc, subject, body, attachmentPath)
    #End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました:" & vbCrLf & Err.Description, vbCritical, "エラー"
End Sub

' ============================================
' Mac用メール作成
' ============================================
Sub CreateEmailMac(recipient As String, cc As String, bcc As String, _
                    subject As String, body As String, attachmentPath As String)
    
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
    mailtoLink = "mailto:" & recipient
    
    ' CCがある場合
    If cc <> "" Then
        mailtoLink = mailtoLink & "?cc=" & EncodeURL(cc)
    End If
    
    ' 件名を追加
    If subject <> "" Then
        If InStr(mailtoLink, "?") > 0 Then
            mailtoLink = mailtoLink & "&subject=" & EncodeURL(subject)
        Else
            mailtoLink = mailtoLink & "?subject=" & EncodeURL(subject)
        End If
    End If
    
    ' 本文を追加（短い場合のみ）
    If Len(body) < 500 Then  ' 長すぎる場合は省略
        If InStr(mailtoLink, "?") > 0 Then
            mailtoLink = mailtoLink & "&body=" & EncodeURL(body)
        Else
            mailtoLink = mailtoLink & "?body=" & EncodeURL(body)
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
        msg = msg & "本文欄に貼り付けてください。" & vbCrLf & vbCrLf
    End If
    
    If bcc <> "" Then
        msg = msg & "注意: BCCはmailto:リンクでは設定できません。" & vbCrLf
        msg = msg & "メールアプリで手動で設定してください。"
    End If
    
    If attachmentPath <> "" Then
        msg = msg & vbCrLf & "注意: 添付ファイルは手動で追加してください。"
    End If
    
    MsgBox msg, vbInformation, "完了"
End Sub

' ============================================
' Windows用メール作成（元のコード）
' ============================================
Sub CreateEmailWindows(recipient As String, cc As String, bcc As String, _
                        subject As String, body As String, attachmentPath As String)
    
    Dim myOL As Object
    Dim myOLMI As Object
    
    ' Outlook を起動（参照設定不要の方法）
    Set myOL = CreateObject("Outlook.Application")
    Set myOLMI = myOL.CreateItem(0)  ' 0 = olMailItem
    
    With myOLMI
        .To = recipient
        If cc <> "" Then .CC = cc
        If bcc <> "" Then .BCC = bcc
        .Subject = subject
        .Body = body
        
        ' 添付ファイルがある場合
        If attachmentPath <> "" Then
            On Error Resume Next
            .Attachments.Add attachmentPath
            On Error GoTo 0
        End If
        
        .Display
    End With
    
    ' クリーンアップ
    Set myOLMI = Nothing
    Set myOL = Nothing
End Sub

' ============================================
' URLエンコード関数
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
