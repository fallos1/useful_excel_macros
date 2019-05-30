' Print page 1 in all worksheets in workbook

Sub printAllFirstPages()

Dim ws As Worksheet
For Each ws In Application.ActiveWorkbook.Worksheets
    ' Print black and white
    ws.PageSetup.BlackAndWhite = True
    ' Print only page 1
    ws.PrintOut from:=1, to:=1
Next
End Sub

' Print between two pages specified by user

Sub printXtoYPages()

Dim ws As Worksheet
Dim xPage as Integer
Dim yPage as Integer
' Messagebox prompt for pages
xPage = InputBox("First Page")
yPage = InputBox("Last Page")
For Each ws In Application.ActiveWorkbook.Worksheets
    ws.PrintOut from:=xPage, to:=yPage
Next
End Sub


