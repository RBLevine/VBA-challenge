Sub clearData()
    Dim ws As Worksheet
    'iterate through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        'delete columns I-P
        For i = 9 To 16:
            ws.Columns(9).EntireColumn.Delete
        Next i
    Next ws
End Sub
