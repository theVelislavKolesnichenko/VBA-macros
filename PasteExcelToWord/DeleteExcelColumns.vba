Sub delCol()
    Dim sourceSheet As Worksheet
    Set sourceSheet = Sheet1
    For i = 2 To 6
        sourceSheet.Columns(i).EntireColumn.Delete
        sourceSheet.Columns(i).EntireColumn.Delete
    Next i
End Sub
