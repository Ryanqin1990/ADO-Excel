# ADO-Excel
这段VBA代码用于快速整理多个结构相同的Excel文件


Dim Filename() As String

Sub 只是个栗子()

    i = MsgBox("请选择文件", vbOKCancel)
    If i <> 1 Then Exit Sub

    Dim I_select As Integer
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = True
        .Show
        If .SelectedItems.Count = 0 Then
            Exit Sub
        Else
            ReDim Filename(1 To .SelectedItems.Count)
            For I_select = 1 To .SelectedItems.Count
                Filename(I_select) = .SelectedItems(I_select)
            Next I_select
        End If
    End With

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
   
    For I_select = 1 To UBound(Filename())
        Set ADO_cn = CreateObject("adodb.connection")
        With ADO_cn
            .Provider = "Microsoft.ACE.OLEDB.12.0"
            .ConnectionString = "Extended Properties = Excel 12.0 macro; Data Source = " & Filename(I_select)
            .Open
        End With
        Sql = "select * from [Sheet1$]"
        Set rs = ADO_cn.Execute(Sql)
        With Sheets("Sheet1")
            For j = 0 To rs.Fields.Count - 1
                .Cells(1, j + 1) = rs.Fields(j).Name
            Next 'j
            .Range("A" & .Cells(1048576, 1).End(xlUp).Row + 1).CopyFromRecordset rs
        End With
        
        ADO_cn.Close
        Set ADO_cn = Nothing
    Next
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    I_path = ThisWorkbook.Path & "\"
    Sheets.Copy
    ActiveWorkbook.SaveAs Filename:=I_path & "自动保存的栗子" & ".xlsx"
    MsgBox ("数据更新完毕，另存在以下目录：" & Chr(10) & I_path & "自动保存的栗子" & ".xlsx")

End Sub
