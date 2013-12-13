
'----------------------------------------------------------
' [Manual Setting]
'99. �J�X�^�� �����⃁�[���{���ŃR�s�[�������Ȃ�ɃJ�X�^�}�C�Y����
'
'----------------------------------------------------------
Private Sub CustomCopyMailRule(objItem As Variant, fldCurrent As Variant)
    ' �R�s�[�i�U�蕪�����[���j �ǉ�
    Dim re As New RegExp
    Set re = New RegExp

       ' f-Base
       '====================================
        re.Pattern = ".*base.*"
        re.IgnoreCase = True
        If re.Test(objItem.Subject) Then
             On Error Resume Next
             objItem.Move fldCurrent.Folders("f-Base")
        ElseIf re.Test(objItem.Body) Then
             On Error Resume Next
             objItem.Move fldCurrent.Folders("f-Base")
        ElseIf re.Test(objItem.To) Then
             On Error Resume Next
             objItem.Move fldCurrent.Folders("f-Base")
        ElseIf re.Test(objItem.CC) Then
             On Error Resume Next
             objItem.Move fldCurrent.Folders("f-Base")
        End If
       '====================================

       ' Calypso
       '====================================
        re.Pattern = ".*calypso.*"
        re.IgnoreCase = True
        If re.Test(objItem.Subject) Then
             On Error Resume Next
             objItem.Move fldCurrent.Folders("calypso")
        ElseIf re.Test(objItem.Body) Then
             On Error Resume Next
             objItem.Move fldCurrent.Folders("calypso")
        ElseIf re.Test(objItem.To) Then
             On Error Resume Next
             objItem.Move fldCurrent.Folders("calypso")
        ElseIf re.Test(objItem.CC) Then
             On Error Resume Next
             objItem.Move fldCurrent.Folders("calypso")
        End If
       '====================================
       
       ' Derico
       '====================================
        re.Pattern = ".*deri.*"
        re.IgnoreCase = True
        If re.Test(objItem.Subject) Then
             On Error Resume Next
             objItem.Move fldCurrent.Folders("DERICO")
        ElseIf re.Test(objItem.Body) Then
             On Error Resume Next
             objItem.Move fldCurrent.Folders("DERICO")
        ElseIf re.Test(objItem.To) Then
             On Error Resume Next
             objItem.Move fldCurrent.Folders("DERICO")
        ElseIf re.Test(objItem.CC) Then
             On Error Resume Next
             objItem.Move fldCurrent.Folders("DERICO")
        End If
       '====================================

       ' �G�N�f��
       '====================================
        re.Pattern = ".*�G�N�f��.*"
        re.IgnoreCase = True
        If re.Test(objItem.Subject) Then
             On Error Resume Next
             objItem.Move fldCurrent.Folders("�G�N�f��")
        ElseIf re.Test(objItem.Body) Then
             On Error Resume Next
             objItem.Move fldCurrent.Folders("�G�N�f��")
        End If
       
        re.Pattern = ".*���[.*"
        If re.Test(objItem.Subject) Then
             On Error Resume Next
             objItem.Move fldCurrent.Folders("�G�N�f��")
        ElseIf re.Test(objItem.Body) Then
             On Error Resume Next
             objItem.Move fldCurrent.Folders("�G�N�f��")
        End If
       '====================================

End Sub
'-----



'----------------------------------------------------------
' [Manual Setting]
'99. �J�X�^�� �F���� & �\���� �������Ȃ�ɃJ�X�^�}�C�Y����
'
'----------------------------------------------------------
Private Sub CustomChangeColorDisplay(objItem As Variant, fldCurrent As Variant)

    If objItem.SenderEmailAddress Like "*@amazon.co.jp*" Then
         objItem.SentOnBehalfOfName = "05Amazon"
         objItem.Save
    End If
    
    If objItem.SenderEmailAddress Like "tsujiiayumi1@yahoo.co.jp" Then
         objItem.Categories = "��"
         objItem.Save
    End If
    
    If objItem.Subject Like "*Base*" Then
         objItem.Categories = "��,��"
         objItem.Save
    End If

End Sub
'-----



'==========================================================
' [Method]
' 00. �蓮 �J�����g�ȉ���Copy���āAPJ�t�H���_�ɃR�s�[�A�Ō��Move
'
'==========================================================
Public Sub A00MyProjectMethod()

    ' �蓮 �J�����g�ȉ���Copy���āAPJ�t�H���_�ɃR�s�[�A�Ō��Move����Method
    
    '[0] �F�ύX
    A01ChangeDisplayColor
    
    Dim i As Integer
    Dim fldCurrent 'As Folder
    Dim objItem 'As MailItem
    Dim objSubFolder 'As Sub Folder
    Dim fldDest 'As Folder
    
    ' ���݂̃t�H���_���擾
    Set fldCurrent = ActiveExplorer.CurrentFolder
    For i = fldCurrent.Items.Count To 1 Step -1
        
        '[1]YYYY�NMM���ɃR�s�[����
        '******************************
        ' ���b�Z�[�W���擾
        Set objItem = fldCurrent.Items(i)
        Set fldDest = fldCurrent.Folders("00_YYYY�NMM��")
        ' Copy����
        Set copyItem = objItem.Copy()
        ' ����
        copyItem.UnRead = False
        copyItem.Save
        On Error Resume Next
        copyItem.Move fldDest
        '******************************
    
        '[2]PJ�Č��̃t�H���_�ɃR�s�[����
        '******************************
        Set copyItem2 = copyItem.Copy()
        CustomCopyMailRule copyItem2, fldCurrent
        '******************************
    
    Next
    
    '[3] �e���M�戶���Move
    A02MoveDirectory
    
End Sub
'-----

'==========================================================
' [Method]
' 01. �蓮 �T�u�t�H���_���܂߂āA�\�����E�F���ނ�ǉ�/�ύX����
'
'==========================================================
Public Sub A01ChangeDisplayColor()

    ' �T�u�t�H���_���܂߂āA�\�����E�F���ނ�ǉ�/�ύX����Method

    Dim i As Integer
    Dim fldCurrent 'As Folder
    Dim objItem 'As MailItem
    Dim objSubFolder 'As Sub Folder
    
    ' ���݂̃t�H���_���擾
    Set fldCurrent = ActiveExplorer.CurrentFolder
    For i = fldCurrent.Items.Count To 1 Step -1
        ' ���b�Z�[�W���擾
        Set objItem = fldCurrent.Items(i)
        MoveOneItemBySender objItem, fldCurrent, "ChangeColor"
    Next

    ' �T�u�t�H���_���擾
    'For Each objSubFolder In fldCurrent.Folders
    '    For i = objSubFolder.Items.Count To 1 Step -1
    '         ' ���b�Z�[�W���擾
    '        Set objItem = objSubFolder.Items(i)
    '        MoveOneItemBySender objItem, objSubFolder, "ChangeColor"
    '    Next
    'Next
End Sub
'-----



'==========================================================
' [Method]
' 02. �蓮 ���̑��̃f�B���N�g����Move����
'
'==========================================================
Public Sub A02MoveDirectory()

    ' ���̑��̃f�B���N�g����Move����Method

    Dim i As Integer
    Dim fldCurrent 'As Folder
    Dim objItem 'As MailItem
    ' ���݂̃t�H���_���擾
    Set fldCurrent = ActiveExplorer.CurrentFolder
    For i = fldCurrent.Items.Count To 1 Step -1
        ' ���b�Z�[�W���擾
        Set objItem = fldCurrent.Items(i)
        MoveOneItemBySender objItem, fldCurrent, "Move"
    Next
End Sub
'-----



'==========================================================
' [Method]
' 02. �蓮 ���̑��̃f�B���N�g����Copy����
'
'==========================================================
Public Sub A02CopyDirectory()

    ' ���̑��̃f�B���N�g����Copy����Method

    Dim i As Integer
    Dim fldCurrent 'As Folder
    Dim objItem 'As MailItem
    ' ���݂̃t�H���_���擾
    Set fldCurrent = ActiveExplorer.CurrentFolder
    For i = fldCurrent.Items.Count To 1 Step -1
        ' ���b�Z�[�W���擾
        Set objItem = fldCurrent.Items(i)
        MoveOneItemBySender objItem, fldCurrent, "Copy"
    Next
End Sub
'-----



'==========================================================
' [Method]
' 03. Auto �����R�s�[�i���[����M���j
'
'==========================================================
Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)

    ' ��M���Ɏ������s�����Method

    Dim i As Integer
    Dim c As Integer
    Dim colID As Variant
    Dim objItem 'As MailItem
    Dim fldCurrent 'As Folder
    
    If InStr(EntryIDCollection, ",") = 0 Then
        Set objItem = Session.GetItemFromID(EntryIDCollection)
        Set fldCurrent = Session.GetDefaultFolder(olFolderInbox)
        'MoveOneItemBySender objItem, fldCurrent, "Copy"
    Else
        colID = Split(EntryIDCollection, ",")
        For i = LBound(colID) To UBound(colID)
        Set objItem = Session.GetItemFromID(colID(i))
        Set fldCurrent = Session.GetDefaultFolder(olFolderInbox)
        'MoveOneItemBySender objItem, fldCurrent, "Copy"
        Next
    End If

End Sub
'-----



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' [Function]
' Function / ���[����Copy����
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CopyMailFolder(objItem As Variant, folderName As String)
    Set fldDest = fldCurrent.Folders(folderName)
    ' Copy����
    Set copyItem = objItem.Copy()
    ' ����
    copyItem.UnRead = False
    copyItem.Save
    On Error Resume Next
    copyItem.Move fldDest
End Sub
'-----



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' [Function]
' ���C��Function / �f�B���N�g�����쐬���āACopy����
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub MoveOneItemBySender(objItem As Variant, fldCurrent As Variant, flagStr As String)

    ' ���C��Function / �f�B���N�g�����쐬���āACopy����

    Dim fldContact 'As Folder
    Dim fldSender 'As Folder
    Dim fldDest 'As Folder
    Dim fldOther 'As Folder
    Dim strFolderName As String
    Dim objContact 'As ContactItem

    ' �A���������
    Set objContact = FindContactByAddress(objItem.SenderEmailAddress, strFolderName)
    Set fldDest = Nothing

    If Not objContact Is Nothing Then
        ' FullName�ɕύX
         objItem.SentOnBehalfOfName = objContact.FullName
         objItem.Save
    End If

    ' �f�B���N�g�����쐬���邩�ǂ����𔻒f����
    '-----------------------------------------------------------------------
    If Not flagStr Like "ChangeColor" Then

        If Not objContact Is Nothing Then
            ' �A�h���X�ŘA���悪���������炻�̃t�H���_�̖��O�̃t�H���_������
            For Each fldContact In fldCurrent.Folders
                If fldContact.Name = strFolderName Then
                    ' �A����t�H���_�Ɠ������O�̃t�H���_������������T�u�t�H���_������
                    For Each fldSender In fldContact.Folders
                        If fldSender.Name = objContact.FullName Then
                            ' �A����̃t���l�[���Ɠ������O�̃t�H���_������������ړ���Ɏw��
                            Set fldDest = fldSender
                        End If
                    Next
                    If fldDest Is Nothing Then
                        ' �ړ���t�H���_��������Ȃ���ΘA����̃t���l�[���Ńt�H���_���쐬���A
                        ' �ړ���t�H���_�Ƃ��Ďw��
                        Set fldDest = fldContact.Folders.Add(objContact.FullName)
                    End If
                End If
            Next
            If fldDest Is Nothing Then
                ' �ړ���t�H���_��������Ȃ���ΘA����t�H���_�̖��O�Ńt�H���_���쐬
                Set fldContact = fldCurrent.Folders.Add(strFolderName)
                ' ����ɘA����̃t���l�[���Ńt�H���_���쐬���A�ړ���t�H���_�Ƃ��Ďw��
                Set fldDest = fldContact.Folders.Add(objContact.FullName)
            End If
            
        Else
            For Each fldOther In fldCurrent.Folders
                If fldOther.Name = "���̑�" Then
                    Set fldDest = fldOther
                End If
            Next
            If fldDest Is Nothing Then
                ' �ړ���t�H���_��������Ȃ���ΘA����t�H���_�̖��O�Ńt�H���_���쐬
                Set fldDest = fldCurrent.Folders.Add("���̑�")
            End If
        End If

    End If
    '-----------------------------------------------------------------------

    ' FullName�ɕύX
    If Not objContact Is Nothing Then
         objItem.SentOnBehalfOfName = objContact.FullName
         objItem.Save
    End If

    ' ------------------------------------
    ' �����Ȃ�̐F���ށ{�\����
    ' ------------------------------------
    CustomChangeColorDisplay objItem, fldCurrent

    ' ------------------------------------
    ' �����Ȃ�̃��[���̐U�蕪��
    ' ------------------------------------
    If flagStr Like "Copy" Then
        CustomCopyMailRule objItem, fldCurrent
    End If
    
    ' ------------------------------------
    ' �ړ���t�H���_�Ƀ��b�Z�[�W��Copy
    ' ------------------------------------
    If flagStr Like "Copy" Then
        ' Copy����
        Set copyItem = objItem.Copy()
        ' ����
        copyItem.UnRead = False
        copyItem.Save
        On Error Resume Next
        copyItem.Move fldDest
    End If

    ' ------------------------------------
    ' �ړ���t�H���_�Ƀ��b�Z�[�W��Move
    ' ------------------------------------
    If flagStr Like "Move" Then
        If Not objItem Is Nothing Then
            ' Move����
            Set copyItem = objItem
            ' ����
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
' ���[���A�h���X����A����̕\�������擾����
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FindContactByAddress(strAddress As String, ByRef strFolderName As String)

' ���[���A�h���X����A����̕\�������擾����Function

    Dim objContacts 'As Folder
    Dim objContact 'As ContactItem
    Dim objSubFolder ' As Folder
    ' ����̘A����t�H���_���擾
    Set objContacts = Application.Session.GetDefaultFolder(olFolderContacts)
    ' ���݂̃t�H���_����ۑ�
    strFolderName = objContacts.Name
    ' �A����t�H���_���ŃA�h���X������
    Set objContact = objContacts.Items.Find("[Email1Address] = '" & strAddress _
        & "' or [Email2Address] = '" & strAddress _
        & "' or [Email3Address] = '" & strAddress & "'")
    If objContact Is Nothing Then
        ' ������Ȃ���΃T�u�t�H���_������
        For Each objSubFolder In objContacts.Folders
            ' ���݂̃t�H���_����ۑ�
            strFolderName = objSubFolder.Name
            ' ���݂̃t�H���_���ŃA�h���X������
            Set objContact = objSubFolder.Items.Find("[Email1Address] = '" & strAddress _
                & "' or [Email2Address] = '" & strAddress _
                & "' or [Email3Address] = '" & strAddress & "'")
            ' ���������烋�[�v���I��
            If Not objContact Is Nothing Then
                Exit For
            End If
        Next
    End If
    Set FindContactByAddress = objContact
End Function
'-----


