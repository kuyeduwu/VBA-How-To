```vb
Sub getTotalRowCount()
  Dim rCount As Long
  Let rCount = ThisWorkbook.Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row
End Sub
```

Actually I have pretty much things to say about the **Total Row Count**, let's start from a simple question: how to locate the **Last Row** when calculating the **Total Row Count**, because you never could count the total rows if you don't know where is the end.

### Locate the Last Row by values in Cell

Basically, start from one cell, if all cells after it are blank 
