'----------------------------------------------------------
' [Manual Setting]
'99. カスタム 件名やメール本文でコピーを自分なりにカスタマイズする
'
'----------------------------------------------------------
Private Sub CustomCopyMailRule(objItem As Variant, fldCurrent As Variant)

' コピー（振り分けルール） 追加
' F-Base
' ==================================
If objItem.Subject Like "*Base*" Then
    Set copyItem = objItem.Copy()
    For Each fldOther In fldCurrent.Folders
        If fldOther.Name = "AAABase" Then
           Set fldDest = fldOther
        End If
    Next

    ' 既読
    copyItem.UnRead = False
    copyItem.Save

    On Error Resume Next
    copyItem.Move fldDest
End If
' ==================================

End Sub
'-----



'----------------------------------------------------------
' [Manual Setting]
'99. カスタム 色分類 & 表示名 を自分なりにカスタマイズする
'
'----------------------------------------------------------
Private Sub CustomChangeColorDisplay(objItem As Variant, fldCurrent As Variant)

    If objItem.SenderEmailAddress Like "*@amazon.co.jp*" Then
         objItem.SentOnBehalfOfName = "05Amazon"
         objItem.Save
    End If
    
    If objItem.SenderEmailAddress Like "tsujiiayumi1@yahoo.co.jp" Then
         objItem.Categories = "青"
         objItem.Save
    End If
    
    If objItem.Subject Like "*Base*" Then
         objItem.Categories = "赤,青"
         objItem.Save
    End If

End Sub
'-----



'==========================================================
' [Method]
' 01. 手動 サブフォルダも含めて、表示名・色分類を追加/変更する
'
'==========================================================
Public Sub A01ChangeDisplayColor()

    ' サブフォルダも含めて、表示名・色分類を追加/変更するMethod

    Dim i As Integer
    Dim fldCurrent 'As Folder
    Dim objItem 'As MailItem
    Dim objSubFolder 'As Sub Folder
    
    ' 現在のフォルダを取得
    Set fldCurrent = ActiveExplorer.CurrentFolder
    For i = fldCurrent.Items.Count To 1 Step -1
        ' メッセージを取得
        Set objItem = fldCurrent.Items(i)
        MoveOneItemBySender objItem, fldCurrent, "ChangeColor"
    Next

    ' サブフォルダを取得
    'For Each objSubFolder In fldCurrent.Folders
    '    For i = objSubFolder.Items.Count To 1 Step -1
    '         ' メッセージを取得
    '        Set objItem = objSubFolder.Items(i)
    '        MoveOneItemBySender objItem, objSubFolder, "ChangeColor"
    '    Next
    'Next
End Sub
'-----



'==========================================================
' [Method]
' 02. 手動 その他のディレクトリをMoveする
'
'==========================================================
Public Sub A02MoveDirectory()

    ' その他のディレクトリをMoveするMethod

    Dim i As Integer
    Dim fldCurrent 'As Folder
    Dim objItem 'As MailItem
    ' 現在のフォルダを取得
    Set fldCurrent = ActiveExplorer.CurrentFolder
    For i = fldCurrent.Items.Count To 1 Step -1
        ' メッセージを取得
        Set objItem = fldCurrent.Items(i)
        MoveOneItemBySender objItem, fldCurrent, "Move"
    Next
End Sub
'-----



'==========================================================
' [Method]
' 02. 手動 その他のディレクトリをCopyする
'
'==========================================================
Public Sub A02CopyDirectory()

    ' その他のディレクトリをCopyするMethod

    Dim i As Integer
    Dim fldCurrent 'As Folder
    Dim objItem 'As MailItem
    ' 現在のフォルダを取得
    Set fldCurrent = ActiveExplorer.CurrentFolder
    For i = fldCurrent.Items.Count To 1 Step -1
        ' メッセージを取得
        Set objItem = fldCurrent.Items(i)
        MoveOneItemBySender objItem, fldCurrent, "Copy"
    Next
End Sub
'-----



'==========================================================
' [Method]
' 03. Auto 自動コピー（メール受信時）
'
'==========================================================
Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)

    ' 受信時に自動実行されるMethod

    Dim i As Integer
    Dim c As Integer
    Dim colID As Variant
    Dim objItem 'As MailItem
    Dim fldCurrent 'As Folder
    
    If InStr(EntryIDCollection, ",") = 0 Then
        Set objItem = Session.GetItemFromID(EntryIDCollection)
        Set fldCurrent = Session.GetDefaultFolder(olFolderInbox)
        MoveOneItemBySender objItem, fldCurrent, "Copy"
    Else
        colID = Split(EntryIDCollection, ",")
        For i = LBound(colID) To UBound(colID)
        Set objItem = Session.GetItemFromID(colID(i))
        Set fldCurrent = Session.GetDefaultFolder(olFolderInbox)
        MoveOneItemBySender objItem, fldCurrent, "Copy"
        Next
    End If

End Sub
'-----



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' [Function]
' メインFunction / ディレクトリを作成して、Copyする
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub MoveOneItemBySender(objItem As Variant, fldCurrent As Variant, flagStr As String)

    ' メインFunction / ディレクトリを作成して、Copyする

    Dim fldContact 'As Folder
    Dim fldSender 'As Folder
    Dim fldDest 'As Folder
    Dim fldOther 'As Folder
    Dim strFolderName As String
    Dim objContact 'As ContactItem

    ' 連絡先を検索    Set objContact = FindContactByAddress(objItem.SenderEmailAddress, strFolderName)
    Set fldDest = Nothing

    If Not objContact Is Nothing Then
        ' FullNameに変更         objItem.SentOnBehalfOfName = objContact.FullName
         objItem.Save
    End If

    ' ディレクトリを作成するかどうかを判断する
    '-----------------------------------------------------------------------
    If Not flagStr Like "ChangeColor" Then

        If Not objContact Is Nothing Then
            ' アドレスで連絡先が見つかったらそのフォルダの名前のフォルダを検索            For Each fldContact In fldCurrent.Folders
                If fldContact.Name = strFolderName Then
                    ' 連絡先フォルダと同じ名前のフォルダが見つかったらサブフォルダを検索                    For Each fldSender In fldContact.Folders
                        If fldSender.Name = objContact.FullName Then
                            ' 連絡先のフルネームと同じ名前のフォルダが見つかったら移動先に指定
                            Set fldDest = fldSender
                        End If
                    Next
                    If fldDest Is Nothing Then
                        ' 移動先フォルダが見つからなければ連絡先のフルネームでフォルダを作成し、
                        ' 移動先フォルダとして指定
                        Set fldDest = fldContact.Folders.Add(objContact.FullName)
                    End If
                End If
            Next
            If fldDest Is Nothing Then
                ' 移動先フォルダが見つからなければ連絡先フォルダの名前でフォルダを作成
                Set fldContact = fldCurrent.Folders.Add(strFolderName)
                ' さらに連絡先のフルネームでフォルダを作成し、移動先フォルダとして指定
                Set fldDest = fldContact.Folders.Add(objContact.FullName)
            End If
            
        Else
            For Each fldOther In fldCurrent.Folders
                If fldOther.Name = "その他" Then
                    Set fldDest = fldOther
                End If
            Next
            If fldDest Is Nothing Then
                ' 移動先フォルダが見つからなければ連絡先フォルダの名前でフォルダを作成
                Set fldDest = fldCurrent.Folders.Add("その他")
            End If
        End If

    End If
    '-----------------------------------------------------------------------

    ' FullNameに変更    If Not objContact Is Nothing Then
         objItem.SentOnBehalfOfName = objContact.FullName
         objItem.Save
    End If

    ' ------------------------------------
    ' 自分なりの色分類＋表示名
    ' ------------------------------------
    CustomChangeColorDisplay objItem, fldCurrent

    ' ------------------------------------
    ' 自分なりのメールの振り分け
    ' ------------------------------------
    If flagStr Like "Copy" Then
        CustomCopyMailRule objItem, fldCurrent
    End If
    
    ' ------------------------------------
    ' 移動先フォルダにメッセージをCopy
    ' ------------------------------------
    If flagStr Like "Copy" Then
        ' Copyする
        Set copyItem = objItem.Copy()
        ' 既読
        copyItem.UnRead = False
        copyItem.Save
        On Error Resume Next
        copyItem.Move fldDest
    End If

    ' ------------------------------------
    ' 移動先フォルダにメッセージをMove
    ' ------------------------------------
    If flagStr Like "Move" Then
        If Not objItem Is Nothing Then
            ' Moveする
            Set copyItem = objItem
            ' 既読
            copyItem.UnRead = False
            copyItem.Save
            
            On Error Resume Next
            copyItem.Move fldDest
        End If
    End If

End Sub
'-----



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' [Function]
' メールアドレスから連絡先の表示名を取得する
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FindContactByAddress(strAddress As String, ByRef strFolderName As String)

' メールアドレスから連絡先の表示名を取得するFunction

    Dim objContacts 'As Folder
    Dim objContact 'As ContactItem
    Dim objSubFolder ' As Folder
    ' 既定の連絡先フォルダを取得
    Set objContacts = Application.Session.GetDefaultFolder(olFolderContacts)
    ' 現在のフォルダ名を保存
    strFolderName = objContacts.Name
    ' 連絡先フォルダ内でアドレスを検索    Set objContact = objContacts.Items.Find("[Email1Address] = '" & strAddress _
        & "' or [Email2Address] = '" & strAddress _
        & "' or [Email3Address] = '" & strAddress & "'")
    If objContact Is Nothing Then
        ' 見つからなければサブフォルダを検索        For Each objSubFolder In objContacts.Folders
            ' 現在のフォルダ名を保存
            strFolderName = objSubFolder.Name
            ' 現在のフォルダ内でアドレスを検索            Set objContact = objSubFolder.Items.Find("[Email1Address] = '" & strAddress _
                & "' or [Email2Address] = '" & strAddress _
                & "' or [Email3Address] = '" & strAddress & "'")
            ' 見つかったらループを終了
            If Not objContact Is Nothing Then
                Exit For
            End If
        Next
    End If
    Set FindContactByAddress = objContact
End Function
'-----



