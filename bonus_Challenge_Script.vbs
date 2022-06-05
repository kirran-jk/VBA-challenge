Attribute VB_Name = "Module3"
Sub bonus_Challenge()

'define variables

Dim ws As Worksheet
Dim LastRow As Long

'runs script on each worksheet

For Each ws In Worksheets

'finds last row of column I

LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

'inputs headings on sheet

With ws

    .Range("P1").Value = "Ticker"
    .Range("Q1").Value = "Value"
    .Range("O2").Value = "Greatest % Increase"
    .Range("O3").Value = "Greatest % Decrease"
    .Range("O4").Value = "Greatest Total Volume"
    
End With

For i = 2 To LastRow

'loops through Percent Changed column to find maximum change

    If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & LastRow)) Then
    ws.Range("P2").Value = ws.Cells(i, 9).Value
    ws.Range("Q2").Value = ws.Cells(i, 11).Value
    ws.Range("Q2").NumberFormat = "0.00%"
    
    End If
    
'loops through Percent Changed column to find minimum change

    If ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & LastRow)) Then
    ws.Range("P3").Value = ws.Cells(i, 9).Value
    ws.Range("Q3").Value = ws.Cells(i, 11).Value
    ws.Range("Q3").NumberFormat = "0.00%"
    
    End If
    
'loops through Total Stock Volume column to find maximum volume

    If ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & LastRow)) Then
    ws.Range("P4").Value = ws.Cells(i, 9).Value
    ws.Range("Q4").Value = ws.Cells(i, 12).Value
    
    End If
    
Next i

Next

End Sub

