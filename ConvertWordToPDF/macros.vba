Sub Print()

    Dim objExcel As Object
    Dim exWb As Object
    
    Set objExcel = CreateObject("Excel.Application")
    
    Set exWb = objExcel.Workbooks.Open("D:\Desktop\Students_20192020.xlsx")
    
    Dim i As Integer
    
    For i = 1 To 389
    
        ActiveDocument.ExportAsFixedFormat OutputFileName:= _
            ("D:\Desktop\BP\" & exWb.Sheets("Sheet1").Cells(i, 8) & ".pdf"), _
            ExportFormat:=wdExportFormatPDF, _
            Range:=wdExportFromTo, From:=i, To:=i
    
    Next i
    
    exWb.Close
    
    Set exWb = Nothing

End Sub
