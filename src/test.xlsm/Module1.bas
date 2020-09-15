Attribute VB_Name = "Module1"
Sub PrintMessage()
'メッセージを表示する

MsgBox , "ようこそgitの世界へ!"

End Sub

Sub readCsv()
'カンマ区切りのcsvファイルを読み込む

    Dim varFileName As Variant
    Dim intFree As Integer
    Dim strRec As String
    Dim strSplit() As String
    Dim i As Long, j As Long
    Dim fileName As String
    Dim ext As String
    Dim delimiterStr As String
    
    
    ' csvファイルの読み込みダイアログからファイルを選択できるように変更
    ChDir ThisWorkbook.Path
    varFileName = Application.GetOpenFilename(FileFilter:="データファイル(*.*),*.*", _
                                                Title:="データファイルの選択")

    ' 拡張子を取得
    ext = Right(varFileName, 3)
    ' 拡張子から分割文字列を設定
    If ext = "csv" Then
        delimiterStr = ","
    ElseIf ext = "tsv" Then
        delimiterStr = Chr(9)
    End If
        
    intFree = FreeFile '空番号を取得
    Open varFileName For Input As #intFree 'CSVファィルをオープン
  
    i = 0
    Do Until EOF(intFree)
        Line Input #intFree, strRec '1行読み込み
        i = i + 1
        strSplit = Split(strRec, delimiterStr) 'カンマ区切りで配列へ
        For j = 0 To UBound(strSplit)
            Cells(i, j + 1) = strSplit(j)
        Next
        '配列をそのまま入れる方法も、ただし全て文字列として入力される
        'Range(Cells(i, 1), Cells(i, UBound(strSplit) + 1)) = strSplit
    Loop
  
    Close #intFree
End Sub


