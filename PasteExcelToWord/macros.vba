Sub Paste()

  Dim objExcel As Object
  Dim exWb As Object

  Set objExcel = CreateObject("Excel.Application")

  Set exWb = objExcel.Workbooks.Open("file path.xlsx")

  Dim i As Integer

  For i = 1 To 389
      Selection.GoTo wdGoToPage, wdGoToAbsolute, i

      Selection.Font.Bold = True
      Selection.Font.Underline = True
      Selection.ParagraphFormat.SpaceAfter = 0
      Selection.ParagraphFormat.SpaceBefore = 12
      Selection.TypeText Text:=exWb.Sheets("Sheet1").Cells(i, 2) & " " & exWb.Sheets("Sheet1").Cells(i, 3)
      Selection.TypeText Text:=MyText
      Selection.InsertParagraph

  Next i


  exWb.Close

  Set exWb = Nothing

End Sub
