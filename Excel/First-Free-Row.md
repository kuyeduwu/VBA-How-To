```vb
Sub getFirstFreeRow()
  Dim rCount As Long
  Let rCount = ThisWorkbook.Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row
End Sub
```
