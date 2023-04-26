Sub ParseInserts()
    Dim fPath As String
    Dim fNum As Integer
    Dim line As String
    Dim values As String
    Dim rowNum As Long
    Dim ws As Worksheet
    
    fPath = "C:\path\to\file.txt"
    rowNum = 1
    
    ' 检查文件是否存在
    If Dir(fPath) = "" Then
        MsgBox "文件不存在！"
        Exit Sub
    End If
    
    ' 打开文件
    fNum = FreeFile
    Open fPath For Input As #fNum
    
    ' 创建新的工作表
    Set ws = ThisWorkbook.Worksheets.Add
    
    ' 读取文件内容并解析
    Do While Not EOF(fNum)
        Line Input #fNum, line ' 逐行读取文件内容
        If InStr(line, "INSERT INTO") > 0 Then ' 找到 INSERT INTO 语句
            values = Mid(line, InStr(line, "VALUES") + 7) ' 提取 VALUES 后面的内容
            values = Replace(values, "'", "") ' 去除单引号
            values = Replace(values, "),(", vbLf) ' 将多个值分隔成多行
            values = Replace(values, "(", "") ' 去除左括号
            values = Replace(values, ")", "") ' 去除右括号
            ' 将多个字段分隔成多列，并写入工作表
            ws.Cells(rowNum, 1).Resize(, UBound(Split(values, ",")) + 1).Value = Split(values, ",")
            rowNum = rowNum + 1 ' 将行号加1
        End If
    Loop
    
    ' 关闭文件
    Close #fNum
    
    ' 设置工作表标题
    ws.Name = Left(fPath, InStrRev(fPath, ".") - 1) ' 工作表名为文件名（不包含扩展名）
    ws.Cells(1, 1).Value = "Column 1"
    ws.Cells(1, 2).Value = "Column 2"
    ' ...
    
    MsgBox "解析完成！"
End Sub
