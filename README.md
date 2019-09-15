# VBA-Challenge
Sub stock()

Dim WS As Worksheet

'Loop throught all worksheets.
For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

Dim RowNumber As Integer
Dim totalVolume As Double
Dim openPrice As Double
Dim closePrice As Double
Dim PercentChange As Double

Cells(1, 9).Value = "Ticker Symbol"
Cells(1, 10).Value = "Price Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Volume"

'Set initial value
totalVolume = 0
RowNumber = 1
openPrice = Cells(2, 3).Value

For i = 2 To LastRow

  If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
  'ticker symbol
    RowNumber = RowNumber + 1
    Cells(RowNumber, 9) = Cells(i, 1).Value
   'price change
    closePrice = Cells(i, 6).Value
    PriceChange = closePrice - openPrice
    Cells(RowNumber, 10).Value = PriceChange
    'percent change
    If openPrice = 0 And closePrice = 0 Then
        PercentChange = 0
    ElseIf openPrice = 0 And closePrice <> 0 Then
        PercentChange = 1
    Else
        PercentChange = PriceChange / openPrice
        Cells(RowNumber, 11).Value = PercentChange
        Cells(RowNumber, 11).NumberFormat = "0.00%"
        End If
    'total volume
    totalVolume = totalVolume + Cells(i, 7).Value
    Cells(RowNumber, 12).Value = totalVolume
    'reset percent change and total volume
    openPrice = Cells(i + 1, 3).Value
    totalVolume = 0
    'total volume if the same ticker
    Else: totalVolume = totalVolume + Cells(i, 7).Value
    
  End If

Next i

'conditional formatting adding color to percentchange

LastRowPC = WS.Cells(Rows.Count, 10).End(xlUp).Row
For ind = 2 To LastRowPC
    If Cells(ind, 10).Value > 0 Or Cells(ind, 10).Value = 0 Then
    Cells(ind, 10).Interior.ColorIndex = 4
    Else
    Cells(ind, 10).Interior.ColorIndex = 3
    End If
    
Next ind

'Challenge question
Cells(2, 14).Value = "Greatest % Increase"
Cells(3, 14).Value = "Greatest % Decrease"
Cells(4, 14).Value = "Greatest total volume"
Cells(1, 15).Value = "Ticker Symbol"
Cells(1, 16).Value = "value"

'Find greatest % increase, greatest % decrease and greatest total volume

maxValue = Application.WorksheetFunction.Max(Columns("K"))
minValue = Application.WorksheetFunction.Min(Columns("K"))
maxVolume = Application.WorksheetFunction.Max(Columns("L"))


For j = 1 To LastRowPC

    If Cells(j, 11).Value = maxValue Then
    Cells(2, 15).Value = Cells(j, 9).Value
    Cells(2, 16).Value = Cells(j, 11).Value
    Cells(2, 16).NumberFormat = "0.00%"
    
    ElseIf Cells(j, 11).Value = minValue Then
    Cells(3, 15).Value = Cells(j, 9).Value
    Cells(3, 16).Value = Cells(j, 11).Value
    Cells(3, 16).NumberFormat = "0.00%"
    
    ElseIf Cells(j, 12).Value = maxVolume Then
    Cells(4, 15).Value = Cells(j, 9).Value
    Cells(4, 16).Value = Cells(j, 12).Value
    End If

    
Next j

Next WS

End Sub



