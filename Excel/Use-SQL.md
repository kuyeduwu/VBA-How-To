```vb
Sub tst()
    Dim Cnn As Object
    Dim Rst As Object
    Dim strCnn As String
    Dim strSql As String
    Dim filePath As String
    
    Set Cnn = CreateObject("ADODB.Connection")
    Set Rst = CreateObject("ADODB.Recordset")
    
    Let filePath = "C:\123.xlsx"
    Let strCnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties=""Excel 12.0;HDR=YES"";"""
    Let strSql = "select * from [Sheet1$]"
    
    Cnn.Open strCnn
    Set Rst = Cnn.Execute(strSql)
    With ThisWorkbook.Sheets(1)
        .Cells.Clear
        For i = 0 To Rst.Fields.Count - 1
            .Cells(1, i + 1) = Rst.Fields(i).Name
        Next
        .Range("A2").CopyFromRecordset Rst
    End With
    Rst.Close
    Cnn.Close
    Set Rst = Nothing
    Set Cnn = Nothing
End Sub
```
