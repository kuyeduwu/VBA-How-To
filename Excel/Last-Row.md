```vb
Sub getLastRow()
  Dim rCount As Long
  Let rCount = ThisWorkbook.Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row
End Sub
```

Actually this is only the simplest and most used option to calculate the last row, and there are something important you should know when using this method:

1. You can only get the last row for 1 column.
  - The above code will get the last row for column A, because it used `.Cells(Rows.Count, 1)`,`1` means column A.
  - If there are more than 1 columns in your table, you should select the column which contains the most rows to calculate the last row for the table.
2. If some rows were hidden at the end of your data, the hidden rows will be ignored, for example:
  - You have 10 rows, row 7, row 8, row 9 and row 10 were hidden, then the last row calculated by the above code will be: 6, this means if you need to write data after the last row by using VBA, you will overwrite some current data.
  - But if some rows between the first row and last row were hidden (e.g.: row 3, row 4, row 5), the above code still can calculate the last row correctly.
3. After the last row, if there is a cell whose value is blank, but there is a formula in it(e.g.: `=""`), the result of the above code will be the row index of that cell.

### Other methods to get the last row
```
Sub getLastRow_Find_Value()
    On Error GoTo err_Handler
    Debug.Print ThisWorkbook.ActiveSheet.Cells.Find("*", SearchOrder:=xlByRows, LookIn:=xlValues, SearchDirection:=xlPrevious).EntireRow.Row
err_Handler:
    Debug.Print "No valid data"
End Sub
```
When using this method:

1. Don't need to care about columns, it will search all columns and give you the maximum index of the last rows.
2. Still, it won't give a correct result if the last row was hidden.
3. For the cells whose value is blank but contain a formula, the above code will ignore those cells, for example:
  - If there are visible values in row 6, also, there are formulas in row 7, but values are blank, then the result of the above code will be: 6
  
```
Sub getLastRow_Find_Formula()
    On Error GoTo err_Handler
    Debug.Print ThisWorkbook.ActiveSheet.Cells.Find("*", SearchOrder:=xlByRows, LookIn:=xlFormulas, SearchDirection:=xlPrevious).EntireRow.Row
err_Handler:
    Debug.Print "No valid data"
End Sub
```
This method is pretty much like the last one, but we changed `LookIn:=xlValues` to `LookIn:=xlFormulas`, this will result in:

1. The code can give correct result even if the last row was hidden.
2. It will not ignore those cells whose value are blank but contain formula.
  - If there are visible values in row 6, also, there are formulas in row 7, but values are blank, then the result of this method will be: 7
  
Also, there are 2 other properties in Excel VBA which can be used to calculate the last row:

  - UsedRange
  - SpecialCells
  
But it is a little complex if you want to get an accurate result, so we won't discuss these 2 properties, go to Google for help if you want to know more.
