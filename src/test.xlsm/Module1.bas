Attribute VB_Name = "Module1"
Sub PrintMessage()
'���b�Z�[�W��\������

MsgBox , "�悤����git�̐��E��!"

End Sub

Sub readCsv()
'�J���}��؂��csv�t�@�C����ǂݍ���

    Dim varFileName As Variant
    Dim intFree As Integer
    Dim strRec As String
    Dim strSplit() As String
    Dim i As Long, j As Long
    Dim fileName As String
    Dim ext As String
    Dim delimiterStr As String
    
    
    ' csv�t�@�C���̓ǂݍ��݃_�C�A���O����t�@�C����I���ł���悤�ɕύX
    ChDir ThisWorkbook.Path
    varFileName = Application.GetOpenFilename(FileFilter:="�f�[�^�t�@�C��(*.*),*.*", _
                                                Title:="�f�[�^�t�@�C���̑I��")

    ' �g���q���擾
    ext = Right(varFileName, 3)
    ' �g���q���番���������ݒ�
    If ext = "csv" Then
        delimiterStr = ","
    ElseIf ext = "tsv" Then
        delimiterStr = Chr(9)
    End If
        
    intFree = FreeFile '��ԍ����擾
    Open varFileName For Input As #intFree 'CSV�t�@�B�����I�[�v��
  
    i = 0
    Do Until EOF(intFree)
        Line Input #intFree, strRec '1�s�ǂݍ���
        i = i + 1
        strSplit = Split(strRec, delimiterStr) '�J���}��؂�Ŕz���
        For j = 0 To UBound(strSplit)
            Cells(i, j + 1) = strSplit(j)
        Next
        '�z������̂܂ܓ������@���A�������S�ĕ�����Ƃ��ē��͂����
        'Range(Cells(i, 1), Cells(i, UBound(strSplit) + 1)) = strSplit
    Loop
  
    Close #intFree
End Sub


