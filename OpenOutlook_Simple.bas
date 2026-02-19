' ============================================
' Outlookを開くシンプルなコード
' Mac/Windows対応版
' ============================================

' ============================================
' 方法1: Outlookアプリケーションを起動（Windows専用）
' ============================================
Sub OpenOutlook_Windows()
    On Error GoTo ErrorHandler
    
    Dim OutApp As Object
    Set OutApp = CreateObject("Outlook.Application")
    
    ' Outlookを表示（既に起動している場合は表示）
    OutApp.Visible = True
    
    MsgBox "Outlookを起動しました。", vbInformation, "完了"
    
    Exit Sub
    
ErrorHandler:
    If Err.Number = 429 Then
        MsgBox "Outlookがインストールされていないか、起動できませんでした。" & vbCrLf & _
               "Mac環境の場合は、OpenOutlook_Mac を使用してください。", vbCritical, "エラー"
    Else
        MsgBox "エラーが発生しました:" & vbCrLf & Err.Description, vbCritical, "エラー"
    End If
End Sub

' ============================================
' 方法2: Mac環境でOutlookを開く
' ============================================
Sub OpenOutlook_Mac()
    On Error GoTo ErrorHandler
    
    ' MacでOutlookアプリケーションを起動
    Shell "open -a ""Microsoft Outlook""", vbHide
    
    If Err.Number <> 0 Then
        ' Outlookが見つからない場合、メールアプリを開く
        Shell "open -a ""Mail""", vbHide
        MsgBox "Outlookが見つかりませんでした。メールアプリを開きました。", vbInformation, "完了"
    Else
        MsgBox "Outlookを起動しました。", vbInformation, "完了"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました:" & vbCrLf & Err.Description, vbCritical, "エラー"
End Sub

' ============================================
' 方法3: Mac/Windows自動判定版（推奨）
' ============================================
Sub OpenOutlook()
    #If Mac Then
        Call OpenOutlook_Mac
    #Else
        Call OpenOutlook_Windows
    #End If
End Sub

' ============================================
' 方法4: メール作成ウィンドウを開く（Windows専用）
' ============================================
Sub OpenNewMail_Windows()
    On Error GoTo ErrorHandler
    
    Dim OutApp As Object
    Dim OutMail As Object
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)  ' 0 = メールアイテム
    
    OutMail.Display
    
    ' クリーンアップ
    Set OutMail = Nothing
    Set OutApp = Nothing
    
    Exit Sub
    
ErrorHandler:
    If Err.Number = 429 Then
        MsgBox "Outlookがインストールされていないか、起動できませんでした。" & vbCrLf & _
               "Mac環境の場合は、OpenNewMail_Mac を使用してください。", vbCritical, "エラー"
    Else
        MsgBox "エラーが発生しました:" & vbCrLf & Err.Description, vbCritical, "エラー"
    End If
End Sub

' ============================================
' 方法5: メール作成ウィンドウを開く（Mac版）
' ============================================
Sub OpenNewMail_Mac()
    On Error GoTo ErrorHandler
    
    ' mailto:リンクでメールアプリを開く
    Shell "open ""mailto:""", vbHide
    
    MsgBox "メールアプリを開きました。", vbInformation, "完了"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました:" & vbCrLf & Err.Description, vbCritical, "エラー"
End Sub

' ============================================
' 方法6: メール作成ウィンドウを開く（Mac/Windows自動判定版）
' ============================================
Sub OpenNewMail()
    #If Mac Then
        Call OpenNewMail_Mac
    #Else
        Call OpenNewMail_Windows
    #End If
End Sub
