' This macro will get the name of all worksheets in the active workbook and paste them in 
' the worksheet "Sheet1". If you want a different sheet change the name. Warning that it
' will replace current data

Sub ListSheets()
  Dim ws As Worksheet
  Dim num as Integer
  
  num = 1
  
  ' Range for paste
  Sheets("Sheet1").Range("A:A").Clear
  
  For Each ws In Worksheets
    Sheets("Sheet1").Cells(num, 1) = ws.Name
    num = num + 1
  Next Ws
  
End Sub
