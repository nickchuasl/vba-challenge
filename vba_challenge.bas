Attribute VB_Name = "vba_challenge"
Sub vba_challenge()

Dim lastRow As Long
Dim stockVolume As Double
Dim countCells As Long
Dim closingValue As Double
Dim openingValue As Double
Dim yearlyChange As Double
Dim percentageChange As Double

For Each ws In Worksheets

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'Create Challenge Summary Table
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"


lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row - 1

For i = 2 To lastRow + 1

If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
stockVolume = stockVolume + ws.Cells(i, 7).Value
countCells = countCells + 1

ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
lastRowSummary = ws.Cells(Rows.Count, 9).End(xlUp).Row
stockVolume = stockVolume + ws.Cells(i, 7).Value

ws.Cells(lastRowSummary + 1, 12).Value = stockVolume

countCells = countCells + 1
closingValue = ws.Cells(i, 6).Value

openingValue = ws.Cells((i - countCells + 1), 6).Value

yearlyChange = closingValue - openingValue

If closingValue <> 0 And openingValue <> 0 Then
percentageChange = (closingValue / openingValue) - 1

Else
ws.Cells(lastRowSummary + 1, 11).Value = 0
ws.Cells(lastRowSummary + 1, 10).Value = 0
End If

countCells = 0
stockVolume = 0

ws.Cells(lastRowSummary + 1, 10).Value = yearlyChange

If closingValue <> 0 And openingValue <> 0 Then
ws.Cells(lastRowSummary + 1, 11).Value = percentageChange
Else
ws.Cells(lastRowSummary + 1, 10).Value = 0
ws.Cells(lastRowSummary + 1, 11).Value = 0

End If

ws.Cells(lastRowSummary + 1, 9).Value = ws.Cells(i, 1).Value


End If
Next i

ws.Range("J:J").Style = "Comma"
'ws.Range("K:K").Style = "Percent"
ws.Range("K:K").NumberFormat = "0.00%"
'ws.Range("L:L").Style = "Comma"
ws.Range("L:L").NumberFormat = "_-* #,##0_-;-* #,##0_-;_-* ""-""??_-;_-@_-"

For j = 2 To lastRowSummary + 1
If ws.Cells(j, 10).Value < 0 Then
ws.Cells(j, 10).Interior.ColorIndex = 3

ElseIf ws.Cells(j, 10).Value >= 0 Then
ws.Cells(j, 10).Interior.ColorIndex = 4


End If
Next j

'Extra Challenge codes
lastRowSummaryComplete = ws.Cells(Rows.Count, 11).End(xlUp).Row

For t = 2 To lastRowSummaryComplete

If ws.Cells(t, 11).Value > 0 Then
    If ws.Cells(t, 11).Value > ws.Cells(t + 1, 11).Value And ws.Cells(t, 11).Value > ws.Range("Q2").Value Then
        ws.Range("Q2").Value = ws.Cells(t, 11).Value
        ws.Range("P2").Value = ws.Cells(t, 9).Value

    End If
    End If

Next t

For s = 2 To lastRowSummaryComplete

If ws.Cells(s, 11).Value < 0 Then
    If ws.Cells(s, 11).Value < ws.Cells(s + 1, 11).Value And ws.Cells(s, 11).Value < ws.Range("Q3").Value Then
        ws.Range("Q3").Value = ws.Cells(s, 11).Value
        ws.Range("P3").Value = ws.Cells(s, 9).Value

    End If
    End If

Next s

'Find the greatest total volume

For u = 2 To lastRowSummaryComplete

If u <> lastRowSummaryComplete Then
    If ws.Cells(u, 12).Value > ws.Cells(u + 1, 12).Value And ws.Cells(u, 12).Value > ws.Range("Q4").Value Then
        ws.Range("Q4").Value = ws.Cells(u, 12).Value
        ws.Range("P4").Value = ws.Cells(u, 9).Value
    End If
ElseIf ws.Cells(u, 12) > ws.Cells(u - 1, 12) Then
    ws.Range("Q4").Value = ws.Cells(u, 12).Value
    ws.Range("P4").Value = ws.Cells(u, 9).Value


    End If

Next u

ws.Range("Q2").Style = "Percent"
'ws.Range("Q3").Style = "Percent"
'ws.Range("Q4").Style = "Comma"
ws.Range("Q3").NumberFormat = "0.00%"
ws.Range("Q4").NumberFormat = "_-* #,##0_-;-* #,##0_-;_-* ""-""??_-;_-@_-"

ws.Columns("I:Q").AutoFit


Next ws
End Sub




