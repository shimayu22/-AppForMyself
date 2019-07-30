Option Explicit

Sub メール内容抽出()


    '今選択しているメールを取得する
    Dim outlookObj As Outlook.Explorer
    Set outlookObj = Application.ActiveExplorer
    Dim outlookSel As Outlook.Selection
    Set outlookSel = outlookObj.Selection
    
    If outlookSel.Count > 1 Or outlookSel.Count < 1 Then
        MsgBox "1つだけ選択してください"
        End
    End If
    
    '取得したメールから各種データを取得する
    Dim outlookItem As MailItem
    Set outlookItem = outlookSel.Item(1)
    
    '送信日時
    Dim transmissionDate As String
    transmissionDate = Format(outlookItem.SentOn, "yyyy/mm/dd hh:nn")
    'タイトル
    Dim subject As String
    subject = outlookItem.subject
    '本文
    Dim body As String
    body = outlookItem.body
    
    'お片付け
    Set outlookItem = Nothing
    Set outlookSel = Nothing
    Set outlookObj = Nothing
    
    
    '取得したデータを編集する
    Dim editTDate As String
    editTDate = "<" & transmissionDate & ">"
    
    Dim kakkoNum As Integer
    kakkoNum = InStr(subject, "]")
    
    Dim editSubject As String
    editSubject = Left(subject, kakkoNum)
    
    Dim pasteData As String
    pasteData = editTDate & vbCrLf & editSubject & vbCrLf & vbCrLf & body
    
    Dim wordCount As Integer
    wordCount = Len(pasteData)
    
    Dim stringLength As Integer
    stringLength = 2000 '2000文字制限
    
    If wordCount > stringLength Then
        pasteData = Left(pasteData, stringLength)
    End If
    
    'クリップボード取得
    With New MSForms.DataObject
        .SetText pasteData
        .PutInClipboard
    End With

End Sub


