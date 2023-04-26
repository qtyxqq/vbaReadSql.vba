Sub ReadTxtFiles()
    ' 设置文件夹路径
    Dim folderPath As String
    folderPath = "差异结果\"
    
    ' 打开Excel工作簿
    Dim wb As Workbook
    Set wb = Workbooks.Add
    
    ' 循环遍历文件夹中的所有txt文件
    Dim file As String
    file = Dir(folderPath & "*.txt")
    Do While Len(file) > 0
        ' 新建工作表并命名
        Dim sheetName As String
        sheetName = Left(file, Len(file) - 4)
        Dim ws As Worksheet
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
        
        ' 读取txt文件中的内容
        Dim content As String
        Open folderPath & file For Input As #1
        content = Input(LOF(1), 1)
        Close #1
        
        ' 分割文本为每个SQL语句
        Dim sqlStatements() As String
        sqlStatements = Split(content, "INSERT")
        Dim statement As Variant
        
        ' 遍历每个SQL语句
        For Each statement In sqlStatements
            ' 如果语句中包含VALUES，说明这是一条INSERT语句
            If InStr(statement, "VALUES") > 0 Then
                ' 截取VALUES后的内容
                Dim values As String
                values = Mid(statement, InStr(statement, "VALUES") + 7)
                values = Left(values, Len(values) - 2) ' 去除末尾的分号和换行符
                
                ' 将VALUES中的字段值分列写入Excel
                Dim valuesArray() As String
                valuesArray = Split(values, ",")
                Dim col As Integer
                col = 1
                Dim value As Variant
                For Each value In valuesArray
                    If InStr(value, """") > 0 Then
                        ' 如果该值中包含双引号，则继续读取直到下一个双引号
                        Dim endPos As Integer
                        endPos = InStr(InStr(value, """") + 1, value, """")
                        value = Mid(value, 2, endPos - 2)
                    End If
                    ws.Cells(1, col).Value = value
                    col = col + 1
                Next value
                ws.Rows(1).EntireColumn.AutoFit ' 自动调整列宽
            End If
        Next statement
        
        file = Dir ' 继续遍历下一个文件
    Loop
End Sub
