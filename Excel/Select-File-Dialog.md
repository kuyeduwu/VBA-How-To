```VB
Sub selectFileDialog()
    Dim wb As Workbook
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Show
        If .SelectedItems.Count = 0 Then
            Exit Sub
        End If
        Set wb = Application.Workbooks.Open(.SelectedItems(1))
    End With
End Sub
```
