Sub ParseTextFiles()
    
    Dim folderPath As String
    Dim fileExtension As String
    Dim targetFile As String
    Dim sqlLine As String
    Dim sqlStatement As String
    Dim fieldValues As Variant
    Dim fieldValue As Variant
    Dim sheetName As String
    Dim rowNum As Long
    Dim colNum As Long
    
    ' 设置文件夹路径和文件扩展名
    folderPath = "差异结果\"
    fileExtension = "*.txt"
    
    ' 打开 Excel 文件
    Dim excelApp As New Excel.Application
    Dim wb As Excel.Workbook
    Dim ws As Excel.Worksheet
    Set wb = excelApp.Workbooks.Add
    Set ws = wb.Worksheets(1)
    
    ' 循环读取每个文件并解析 INSERT 语句中的字段值
    targetFile = Dir(folderPath & fileExtension)
    Do While targetFile <> ""
        ' 获取文件名作为表格名
        sheetName = Left(targetFile, InStrRev(targetFile, ".") - 1)
        ' 新建工作表并命名
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
        ' 初始化行号
        rowNum = 1
        ' 打开文本文件并逐行解析 INSERT 语句中的字段值
        Open folderPath & targetFile For Input As #1
        Do Until EOF(1)
            Line Input #1, sqlLine
            ' 检查行是否包含 INSERT 语句
            If InStr(1, sqlLine, "INSERT INTO", vbTextCompare) > 0 Then
                ' 解析 INSERT 语句并获取字段值
                sqlStatement = Replace(sqlLine, "INSERT INTO", "VALUES", , , vbTextCompare)
                sqlStatement = Replace(sqlStatement, ";", "", , , vbTextCompare)
                sqlStatement = Replace(sqlStatement, "(", "", , , vbTextCompare)
                sqlStatement = Replace(sqlStatement, ")", "", , , vbTextCompare)
                sqlStatement = Replace(sqlStatement, ",", "|", , , vbTextCompare)
                fieldValues = Split(sqlStatement, "|")
                ' 将字段值写入表格
                colNum = 1
                For Each fieldValue In fieldValues
                    ws.Cells(rowNum, colNum) = fieldValue
                    colNum = colNum + 1
                Next fieldValue
                rowNum = rowNum + 1
            End If
        Loop
        Close #1
        ' 移动到下一个文件
        targetFile = Dir()
    Loop
    
    ' 保存 Excel 文件并退出
    wb.SaveAs ThisWorkbook.Path & "\parsed_data.xlsx", FileFormat:=51
    wb.Close
    excelApp.Quit
    
    MsgBox "解析完成！"
    
End Sub
