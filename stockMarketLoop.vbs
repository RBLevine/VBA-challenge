Sub stockMarketLoop()
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
    'define variables
        'variable to store last row number
        Dim lastRow As Double
        lastRow = Range("A1").End(xlDown).Row
        'MsgBox ("Last Row: " & lastRow)
        
        'variable for yearly change and percent change
        Dim yearlyChange As Double
        Dim percentChange As Double
            
        'variable to iterate through rows
        Dim i As Double
        
        'variables to keep to track of number of tickers
        Dim count As Integer
        count = 2
        'MsgBox ("Count: " & count)
            
        'variables to store first open amount, last closing amount for each ticker
        Dim openAmount As Double
        openAmount = ws.Cells(2, 3).Value
        Dim closeAmount As Double
            
        'variables to store total vol for each ticker, set to 0
        Dim volTotal As Double
        volTotal = 0
        'MsgBox ("Volume total: " & volTotal)
        
        'variable to keep track of which row to print the data on
        Dim printRow As Double
        
        'print headers for output row
       ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
            
        For i = 2 To lastRow:
           ' printRow = Range("I1").CurrentRegion.End(xlDown).Row + 1
            
            If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
                volTotal = ws.Cells(i, 7) + volTotal
            
            ElseIf ws.Cells(i + 1).Value <> ws.Cells(i, 1).Value Then
                'if current cell has a new ticker: record data to spreadsheet
                volTotal = ws.Cells(i, 7) + volTotal
                'record the closing amount
                closeAmount = ws.Cells(i, 6).Value
                'print the ticker amount
                ws.Cells(count, 9).Value = ws.Cells(i, 1).Value
                'print the yearly change amount
                yearlyChange = closeAmount - openAmount
                ws.Cells(count, 10).Value = yearlyChange
                'conditional formatting for positive and negative change
                If ws.Cells(count, 10).Value < 0 Then
                    ws.Cells(count, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(count, 10).Interior.ColorIndex = 4
                End If
                'print the percent change and handle div by zero error
                If yearlyChange = 0 Then
                    percentChange = 0
                    ws.Cells(count, 11).Value = percentChange
                ElseIf openAmount = 0 And closeAmount <> 0 Then
                    percentChange = -1
                    ws.Cells(count, 11).Value = percentChange
                Else
                    percentChange = yearlyChange / openAmount
                    ws.Cells(count, 11).Value = percentChange
                    ws.Cells(count, 11) = Format(ws.Cells(count, 11).Value, "Percent")
                End If
            
                
                ws.Cells(count, 12).Value = volTotal
                
                'reset data with new ticker info
                count = count + 1
                openAmount = ws.Cells(i + 1, 3).Value
                volTotal = 0
            End If
        Next i
        
        
        'find last row of the printed data
        Dim lastRow2 As Double
        lastRow2 = ws.Range("I1").CurrentRegion.End(xlDown).Row
        
        'set variables for finding greatest increase, decrease, and volume
        'create place holder for max percent amount, and ticker, set value to 0
        Dim MaxValue As Double
        Dim maxTic As String
        MaxValue = 0
        
        'create place holder for min percent amount, and ticker, set value to 0
        Dim MinValue As Double
        Dim minTic As String
        MinValue = 0
        
        'create place holder for max volume amount, and ticker, set value to 0
        Dim GreatestVolume As Double
        Dim volTic As String
        GreatestVolume = 0
        
        'goes through printed, condensed data
        For i = 2 To lastRow2:
            If ws.Cells(i, 11).Value <> 0 Then
                'if percent is greater than current MaxValue, set value to new percent, update ticker
                If ws.Cells(i, 11).Value > MaxValue Then
                    MaxValue = ws.Cells(i, 11).Value
                    maxTic = ws.Cells(i, 9).Value
                'if percent is less than current MinValue, set value to new percent, update ticker
                ElseIf ws.Cells(i, 11) < MinValue Then
                    MinValue = ws.Cells(i, 11).Value
                    minTic = ws.Cells(i, 9).Value
                End If
            End If
            
            If ws.Cells(i, 12) <> 0 Then
                'if volume is greater than current GreatestVolume, set value to new percent, update ticker
                If ws.Cells(i, 12) > GreatestVolume Then
                    GreatestVolume = ws.Cells(i, 12)
                    volTic = ws.Cells(i, 9).Value
                End If
            End If
        Next i
        
        'print out headers of columns and rows
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        
        'print out max value and ticker, format to percent
        ws.Cells(2, 15).Value = maxTic
        ws.Cells(2, 16).Value = MaxValue
        ws.Cells(2, 16) = Format(ws.Cells(2, 16).Value, "Percent")
        
        'print out min value and ticker, format to percent
        ws.Cells(3, 15).Value = minTic
        ws.Cells(3, 16).Value = MinValue
        ws.Cells(3, 16) = Format(ws.Cells(3, 16).Value, "Percent")
        
        'print out volume and ticker
        ws.Cells(4, 15).Value = volTic
        ws.Cells(4, 16).Value = GreatestVolume
    
    'set entire used range to AutoFit so all data can easily be seen
    ws.UsedRange.Columns.AutoFit
    
    
Next ws

End Sub



