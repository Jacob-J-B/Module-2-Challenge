Attribute VB_Name = "Module1"
Sub Stocks()

'create a loop to run on all worksheets
Dim ws As Worksheet
For Each ws In Worksheets

'declare variables used for calculations and to hold results
Dim i As Long
Dim StockSy As String
Dim RowLoc As Integer
Dim StartPrice As Double
Dim EndPrice As Double
Dim TotalVol As Double
TotalVol = 0
RowLoc = 2
Dim FirstPass As Boolean
FirstPass = True
StartPrice = 0

'create row headers for results, format row width
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Quarterly Change"
ws.Columns("J").ColumnWidth = 17
ws.Cells(1, 11).Value = "Percent Change"
ws.Columns("K").ColumnWidth = 15
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Columns("L").ColumnWidth = 18

'determine length of ticker row, hold as TickerRow variable
Dim TickerRow As Long
TickerRow = Cells(Rows.Count, 1).End(xlUp).Row

'loop through all rows
For i = 2 To TickerRow

'there are three conditions for the following If statement:
'the last time looking at a given symbol, the first time, and all other times
'first, if the ticker symbol changes it is the last time through:
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
'the ticker symbol is displayed in results
    StockSy = ws.Cells(i, 1).Value
    ws.Range("I" & RowLoc).Value = StockSy
'the last day closing price is captured
    EndPrice = Cells(i, 6).Value
'the total volume is calculated and displayed in results
    TotalVol = TotalVol + ws.Cells(i, 7).Value
    ws.Cells(RowLoc, 12).Value = TotalVol
'the quarter change is calculated and displated in results
    ws.Cells(RowLoc, 10).Value = (EndPrice - StartPrice)
'the quarter change is calculated as a percent and displated in results
    ws.Cells(RowLoc, 11).Value = (ws.Cells(RowLoc, 10).Value / StartPrice)
    ws.Cells(RowLoc, 11).NumberFormat = "0.00%"
'the quarter change and percent change is formatted, red for negative, green for positive
    If ws.Cells(RowLoc, 11).Value > 0 Then
    ws.Cells(RowLoc, 11).Interior.ColorIndex = 4
    ElseIf ws.Cells(RowLoc, 11).Value < 0 Then
    ws.Cells(RowLoc, 11).Interior.ColorIndex = 3
    End If
    If ws.Cells(RowLoc, 10).Value > 0 Then
    ws.Cells(RowLoc, 10).Interior.ColorIndex = 4
    ElseIf ws.Cells(RowLoc, 10).Value < 0 Then
    ws.Cells(RowLoc, 10).Interior.ColorIndex = 3
    End If
'RowLoc is updated so next time the results are printed on the next row
    RowLoc = RowLoc + 1
'the total volume counter variable TotalVol is reset for the next symbol
    TotalVol = 0
'the FirstPass varbiable is reset to true for the next symbol
    FirstPass = True
'next is the first look at a given symbol, the total volume variable is reset
'and the tally of total volume via the TotalVol variable starts
    ElseIf FirstPass = True Then
    TotalVol = 0
    TotalVol = ws.Cells(i, 7).Value
'the starting price is captured
    StartPrice = ws.Cells(i, 3).Value
'the FirstPass variable is set to false
    FirstPass = False
'for all other iterations through the loop the total volume tally increased
    Else
    TotalVol = TotalVol + ws.Cells(i, 7).Value
    End If

Next i

'create row/colum headers for %inc,%dec,greatest vol
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"
ws.Columns("N").ColumnWidth = 20

'create variables used when calculating greatest inc, dec, vol values and names
Dim GreatInc As Double
Dim GreatDec As Double
Dim GreatVol As Double
GreatInc = 0
GreatDec = 0
GreatVol = 0
Dim GreatIncSym As String
GreatIncSym = ""
Dim GreatDecSym As String
Dim GreatVolSym As String

'determine length of results row, hold as ResultsRow
Dim ResultsRow As Long
ResultsRow = Cells(Rows.Count, 9).End(xlUp).Row

'loop through results, find greatest increase, store name+value
For j = 2 To ResultsRow
    If ws.Cells(j, 11).Value > GreatInc Then
    GreatInc = ws.Cells(j, 11).Value
    GreatIncSym = ws.Cells(j, 9).Value
    End If
Next j
'print and format greatest increase
ws.Cells(2, 16).Value = GreatInc
ws.Cells(2, 16).NumberFormat = "0.00%"
ws.Cells(2, 15).Value = GreatIncSym

'loop through results, find greatest decrease, store name+value
For j = 2 To ResultsRow
    If ws.Cells(j, 11).Value < GreatDec Then
    GreatDec = ws.Cells(j, 11).Value
    GreatDecSym = ws.Cells(j, 9).Value
    End If
Next j
'print and format greatest decrease
ws.Cells(3, 16).Value = GreatDec
ws.Cells(3, 16).NumberFormat = "0.00%"
ws.Cells(3, 15).Value = GreatDecSym

'loop through results, find greatest volume, store name+value
For j = 2 To ResultsRow
    If ws.Cells(j, 12).Value > GreatVol Then
    GreatVol = ws.Cells(j, 12).Value
    GreatVolSym = ws.Cells(j, 9).Value
    End If
Next j
'print and greatest decrease
ws.Cells(4, 16).Value = GreatVol
ws.Cells(4, 15).Value = GreatVolSym

Next ws

End Sub
