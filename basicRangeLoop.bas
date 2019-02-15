'A basic range loop that populates the current selection
Sub CellLoop()

Dim cell As Range

counter = 1

For Each cell In Selection
  cell = counter
  counter = counter + 1
Next

End Sub