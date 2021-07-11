Sub FindandReplaceText()

    Dim specialcharacters As Variant
    specialcharacters = Array("~", "=", "#", "{", "}", ":", Chr(10))
    Dim replaceWith As Variant
    replaceWith = Array("\~", "\=", "\#", "\{", "\}", "\:", "\n")

    Dim xFind As String
    Dim xRep As String
    Dim xRg As Range
    On Error Resume Next
    Set xRg = Selection
    
    For i = 1 To UBound(specialcharacters)
    
        xFind = specialcharacters(i)
        xRep = replaceWith(i)
        If xFind = "False" Or xRep = "False" Then Exit Sub
        xRg.Replace xFind, xRep, xlPart, xlByRows, False, False, False, False
    
    Next i
    
End Sub
