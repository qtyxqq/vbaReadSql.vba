Sub ReadTxtFilesAndExportToExcel()
    '现文件路径和名称
    Dim oldFilePath As String
    '使用文件选择器让用户选择现文件
    oldFilePath = Application.GetOpenFilename("Text Files (*.txt), *.txt")
    '如果用户没有选择文件则退出
    If oldFilePath = False Then Exit Sub
    
    '新文件路径和名称
    Dim newFilePath As String
    '使用文件选择器让用户选择新文件
    newFilePath = Application.GetOpenFilename("Text Files (*.txt), *.txt")
    '如果用户没有选择文件则退出
    If newFilePath = False Then Exit Sub
    
    '打开“现”工作表
    Sheets("现").Select
    '从B13单元格开始输出
    Dim currentRow As Long
    currentRow = 13
    '读取“现”文件
    Dim oldFile As Integer
    oldFile = FreeFile()
    Open oldFilePath For Input As #oldFile
    While Not EOF(oldFile)
        Line Input #oldFile, textline
        If InStr(1, textline, vbTab) > 0 Then
            Dim textArray() As String
            textArray = Split(textline, vbTab)
            Dim i As Integer
            For i = 0 To UBound(textArray)
                Cells(currentRow, 2 + i).Value = textArray(i)
            Next i
        Else
            Cells(currentRow, 2).Value = textline
        End If
        currentRow = currentRow + 1
    Wend
    Close #oldFile
    
    '打开“新”工作表
    Sheets("新").Select
    '从B13单元格开始输出
    currentRow = 13
    '读取“新”文件
    Dim newFile As Integer
    newFile = FreeFile()
    Open newFilePath For Input As #newFile
    While Not EOF(newFile)
        Line Input #newFile, textline
        If InStr(1, textline, vbTab) > 0 Then
            Dim textArray() As String
            textArray = Split(textline, vbTab)
            Dim i As Integer
            For i = 0 To UBound(textArray)
                Cells(currentRow, 2 + i).Value = textArray(i)
            Next i
        Else
            Cells(currentRow, 2).Value = textline
        End If
        currentRow = currentRow + 1
    Wend
    Close #newFile
End Sub