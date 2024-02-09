Attribute VB_Name = "Module1"
Sub MultipleYearStockData():

'Assign Variables'
For Each ws In Worksheets
Dim Worksheet As String
Dim i As Long
Dim j As Long
Dim TickerRow As Long
Dim LastRowIndexA As Long
Dim PercentChange As Double
    
    
    
'Make columns'
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
'Complete Ticker Data'
TickerRow = 2
j = 2

'Find how many rows contain data in a ws that contains data in column "A"'
LastRowIndexA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
For i = 2 To LastRowIndexA
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    ws.Cells(TickerRow, 9).Value = ws.Cells(i, 1).Value
  
'Yearly Change Data.'
    ws.Cells(TickerRow, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value

'Conditional Formatting'
    If ws.Cells(TickerRow, 10).Value < 0 Then
        ws.Cells(TickerRow, 10).Interior.ColorIndex = 3 'red = negative'
        Else
        ws.Cells(TickerRow, 10).Interior.ColorIndex = 4 'green = positive'
        End If
        
'Percent Change'
    If ws.Cells(j, 3).Value <> 0 Then
    PercentChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
    ws.Cells(TickerRow, 11).Value = Format(PercentChange, "Percent")
    Else
    ws.Cells(TickerRow, 11).Value = Format(0, "Percent")
    End If

    ws.Cells(TickerRow, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
        TickerRow = TickerRow + 1
        j = i + 1
                
        End If
    Next i





'"Greatest % increase", "Greatest % decrease", and "Greatest total volume"'
Dim LastRowIndexI As Long
Dim PercentInc As Double
PercentInc = ws.Cells(2, 11).Value

Dim PercentDec As Double
PercentDec = ws.Cells(2, 11).Value

Dim TotalVol As Double
TotalVol = ws.Cells(2, 12).Value


    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
'Find how many rows contain data in a ws that contains data in column "I"'
LastRowIndexI = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row

For i = 2 To LastRowIndexI
    If ws.Cells(i, 12).Value > TotalVol Then
    TotalVol = ws.Cells(i, 12).Value
    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    Else
    TotalVol = TotalVol
    End If
    
    If ws.Cells(i, 11).Value > PercentInc Then
    PercentInc = ws.Cells(i, 11).Value
    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
    Else
    PercentInc = PercentInc
    End If
    
    If ws.Cells(i, 11).Value < PercentDec Then
    PercentDec = ws.Cells(i, 11).Value
    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
    Else
    PercentDec = PercentDec
    End If
    
'Format to Percent'
ws.Cells(2, 17).Value = Format(PercentInc, "Percent")
ws.Cells(3, 17).Value = Format(PercentDec, "Percent")
ws.Cells(4, 17).Value = Format(TotalVol, "Percent")

Next i
             

   
Next ws

End Sub
