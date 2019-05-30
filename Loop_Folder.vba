' Loop through all files in a folder with specific extention

Sub LoopFolder()
    Dim folderPath As String
    Dim fileName As String
    Dim fileExtention as String
    Dim wb As Workbook
  
    folderPath = "C:\Users\test"
    ' Loop only specific extention. This case ".xlsx" (excel)
    fileExtention = "xlsx"
    
    ' Add "\" to folderPath If last "\" left out
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath + "\"
    

    fileName = Dir(folderPath & "*." & fileExtention)
    Do While filename <> ""
      ' Set ScreenUpdating to TRUE if debugging
      Application.ScreenUpdating = False
        Set wb = Workbooks.Open(folderPath & filename)
        ' Code to run on each file
        Call AnotherMacro
        
        wb.Save
        wb.Close
        filename = Dir
    Loop
  Application.ScreenUpdating = True

End Sub
