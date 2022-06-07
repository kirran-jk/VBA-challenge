Attribute VB_Name = "Module2"
Sub vba_Challenge()

'define variables

Dim ws As Worksheet
Dim lastRow As Long
Dim tickerName As String
Dim openValue As Double
Dim closeValue As Double
Dim stockTotal As Double
Dim summaryRow As Integer

'run script on each worksheet in the workbook

For Each ws In Worksheets

    'find last row of data

    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'set initial variable for total stock volume per ticker

    stockTotal = 0
    
    'set initial row location for ticker in summary table

    summaryRow = 2

    'print headings for summary table on sheet

    With ws
    
        .Range("I1").Value = "Ticker"
        .Range("J1").Value = "Yearly Change"
        .Range("K1").Value = "Percent Change"
        .Range("L1").Value = "Total Stock Volume"
    
    End With
    
    'loop through all stock data
    
    For i = 2 To lastRow
    
        'check if the ticker on the current row is equal to the row below, if not take the ticker name, the closing value for the year, and add to the total stock volume
        
        If ws.Cells(i, 1).Value <> ws.Cells((i + 1), 1).Value Then
         tickerName = ws.Cells(i, 1).Value
         closeValue = ws.Cells(i, 6).Value
         stockTotal = stockTotal + ws.Cells(i, 7).Value
         
         'in the summary table: print the name, calculate yearly difference, calculate the % change in the year, and print the total stock volume
        
         ws.Cells(summaryRow, 9).Value = tickerName
         ws.Cells(summaryRow, 10).Value = closeValue - openValue
         ws.Cells(summaryRow, 10).NumberFormat = "0.00"
         ws.Cells(summaryRow, 11).Value = (closeValue - openValue) / openValue
         ws.Cells(summaryRow, 11).NumberFormat = "0.00%"
         ws.Cells(summaryRow, 12).Value = stockTotal
            
            'format the yearly difference cells; red for negative values, green for positive values, and grey for zeros
            
            If ws.Cells(summaryRow, 10).Value < 0 Then
             ws.Cells(summaryRow, 10).Interior.ColorIndex = 3
            
            ElseIf ws.Cells(summaryRow, 10).Value > 0 Then
             ws.Cells(summaryRow, 10).Interior.ColorIndex = 4
            
            Else
             ws.Cells(summaryRow, 10).Interior.ColorIndex = 15
             
            End If
        
        'add one to the summary table row
        
        summaryRow = summaryRow + 1
        
        'reset stock total to zero for the next ticker
        
        stockTotal = 0
        
        'if the loop is still within one ticker, add the volume to the total stock volume
        
        Else
         stockTotal = stockTotal + ws.Cells(i, 7).Value
        
        End If
        
        'check if ticker on the row above is equal to the current row, if so take the opening value for the beginning of the year for the ticker
        
        If ws.Cells(i, 1).Value <> ws.Cells((i - 1), 1).Value Then
         openValue = ws.Cells(i, 3).Value
         
        End If
        
    Next i

Next
    
End Sub
