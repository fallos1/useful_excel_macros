' Convert all numbers formatted as text to numbers
' If a range is selected, only convert items in range

Sub TextToNumber()

  Dim xCell As Range
    ' If there is at least one selection
    If Selection.Count > 1 Then
        For Each xCell In Selection
            If IsNumeric(xCell) And xCell <> "" Then c.Value = Val(xCell.Value)
        Next
    ' If no selection convert all in active worksheet
    Else
    'IF NO SELECTION IS MADE, THEN CONVERT EVERY CELL WITHIN THE USED RANGE
        For Each xRange In ActiveSheet.UsedRange
            If IsNumeric(xCell) And xCell <> "" Then xCell.Value = Val(xCell.Value)
        Next

    End If


End Sub
