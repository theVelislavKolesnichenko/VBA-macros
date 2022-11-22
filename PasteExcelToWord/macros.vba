Sub setStartPage(page As Integer)
'
' Macro1 Macro
'
'
    If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
    End If
    If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
        ActivePane.View.Type = wdOutlineView Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    With Selection.HeaderFooter.PageNumbers
        .NumberStyle = wdPageNumberStyleArabic
        .HeadingLevelForChapter = 0
        .IncludeChapterNumber = False
        .ChapterPageSeparator = wdSeparatorHyphen
        .RestartNumberingAtSection = True
        .StartingNumber = page
    End With
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
End Sub

Sub Paste()

  Dim objExcel As Object
  Dim exWb As Object

  Set objExcel = CreateObject("Excel.Application")

  Set exWb = objExcel.Workbooks.Open("D:\OneDrive - tu-varna.bg\BasicProgramming\2022_2023\СДР\Students.xlsx")

  Dim i As Integer

  For i = 1 To 40
      Selection.GoTo wdGoToPage, wdGoToAbsolute, i

    If i Mod 21 = 0 Then
        setStartPage (i)
    End If


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
