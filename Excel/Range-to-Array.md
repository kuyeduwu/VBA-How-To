```vb
Sub Range2Ary()
    Dim arr
    With ThisWorkbook.ActiveSheet
        Let arr = .Range(.Cells(1, 1), .Cells(10, 2)).Value
    End With
End Sub
```

The `arr` will be 2-d array, even when there is only 1 row / column in your range, that means: if you want to get the first data in `arr`, you cannot use `arr(1)`, you must use `arr(1, 1)`.
