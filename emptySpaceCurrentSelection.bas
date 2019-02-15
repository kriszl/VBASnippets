'In a selection range determine if there are any empty cells.
Sub CellLoop()

countEmptyCells = 0

For Each cell In Selection
    Dim isMyCellEmpty As Boolean
    isMyCellEmpty = IsEmpty(cell)
    
    If isMyCellEmpty = True Then
        countEmptyCells = countEmptyCells + 1
    End If
Next

If countEmptyCells = 0 Then
    MsgBox "There are no empty cells in the selection."
Else
    MsgBox "There are empty cells in the selection."
End If

End Sub
