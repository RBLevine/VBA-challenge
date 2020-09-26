Sub clearData()
    Dim ws As Worksheet
    'iterate through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        'delete columns I-P
        For i = 9 To 16:
            Columns(i).EntireColumn.Delete
        Next i
    Next ws
End Sub
