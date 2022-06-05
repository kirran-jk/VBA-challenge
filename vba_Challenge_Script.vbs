Attribute VB_Name = "Module2"
Sub vba_Challenge()

' define variables

Dim ws As Worksheet
Dim LastRow As Long
Dim OpenValue As Double
Dim VolStart As String
Dim VolEnd As String
Dim TickerRow As Integer

'runs script on each worksheet

For Each ws In Worksheets

'calculates last row of data on sheet and defines row where values will be input

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
TickerRow = 2

'inputs headings on sheet

With ws

    .Range("I1").Value = "Ticker"
    .Range("J1").Value = "Yearly Change"
    .Range("K1").Value = "Percent Change"
    .Range("L1").Value = "Total Stock Volume"

End With

'loop to find unique ticker names, then inputs into column I
'then calculates Yearly Change; the difference between the opening value at the beginning of the year and the closing value at the end of the year.

For i = 2 To LastRow

    If ws.Cells(i, 1).Value <> ws.Cells((i + 1), 1).Value Then
    ws.Cells(TickerRow, 9).Value = ws.Cells(i, 1).Value
    ws.Cells(TickerRow, 10).Value = ws.Cells(i, 6).Value - OpenValue
    ws.Cells(TickerRow, 10).NumberFormat = "0.00"

'formats Yearly Change column; green if positive, red if negative, and grey if 0

        If ws.Cells(TickerRow, 10).Value < 0 Then
        ws.Cells(TickerRow, 10).Interior.ColorIndex = 3
        
        ElseIf ws.Cells(TickerRow, 10).Value > 0 Then
        ws.Cells(TickerRow, 10).Interior.ColorIndex = 4
        
        Else
        ws.Cells(TickerRow, 10).Interior.ColorIndex = 15
        
        End If

'calculates Percent Change; Yearly Change / opening value
'finds the last volume cell of the year and sums the volume of stock for the year
'adds 1 to TickerRow in order to input the next ticker on the row below

    ws.Cells(TickerRow, 11).Value = ws.Cells(TickerRow, 10).Value / OpenValue
    ws.Cells(TickerRow, 11).NumberFormat = "0.00%"
    VolEnd = "G" & i
    ws.Cells(TickerRow, 12).Value = Application.WorksheetFunction.Sum(ws.Range(VolStart & ":" & VolEnd))
    TickerRow = TickerRow + 1

    End If
    
'stores the opening value and cell of the opening volume of the ticker, until another ticker is found

    If ws.Cells(i, 1).Value <> ws.Cells((i - 1), 1).Value Then
    OpenValue = ws.Cells(i, 3).Value
    VolStart = "G" & i
    
    End If
    
Next i

Next

End Sub
