Attribute VB_Name = "Module1"
Sub StockInfo()

For Each ws In Worksheets

ws.Range("I1").Value = "Ticker Symbol"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percentage Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"
ws.Range("O1").Value = "Ticker Symbol"
ws.Range("P1").Value = "Value"

Dim SummaryTable As Integer
SummaryTable = 2
Dim TickerSymbol As String
Dim TotalVolume As LongLong
TotalVolume = 0
Dim YearlyOpen As Double
YearlyOpen = ws.Cells(2, 3).Value
Dim YearlyClose As Double
YearlyClose = 0
Dim PercentChange As Double
Dim YearlyChange As Double

TotalRows = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To TotalRows

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

TickerSymbol = ws.Cells(i, 1).Value
TotalVolume = TotalVolume + ws.Cells(i, 7).Value
YearlyClose = ws.Cells(i, 6).Value
YearlyChange = (YearlyClose / YearlyOpen) - 1

ws.Range("J" & SummaryTable).Value = YearlyClose - YearlyOpen
ws.Range("K" & SummaryTable).Value = YearlyChange
ws.Range("I" & SummaryTable).Value = TickerSymbol
ws.Range("L" & SummaryTable).Value = TotalVolume

SummaryTable = SummaryTable + 1
TotalVolume = 0
YearlyOpen = ws.Cells(i + 1, 3).Value

Else

TotalVolume = TotalVolume + ws.Cells(i, 7).Value

End If

Next i

TotalSummaryRows = ws.Cells(Rows.Count, 9).End(xlUp).Row
Dim GreatChange As Double
GreatChange = 0
GreatDecrease = 0
GreatTotalVolume = 0
Dim TickerGreat As String
Dim TickerLow As String
Dim TickerVol As String

For x = 2 To TotalSummaryRows

If ws.Cells(x, 11).Value > GreatChange Then
GreatChange = ws.Cells(x, 11).Value
TickerGreat = ws.Cells(x, 9).Value
ws.Range("P2").Value = GreatChange
ws.Range("O2").Value = TickerGreat

End If

If ws.Cells(x, 11).Value < GreatDecrease Then
GreatDecrease = ws.Cells(x, 11).Value
TickerLow = ws.Cells(x, 9).Value
ws.Range("P3").Value = GreatDecrease
ws.Range("O3").Value = TickerLow

End If

If ws.Cells(x, 12).Value > GreatTotalVolume Then
GreatTotalVolume = ws.Cells(x, 12).Value
TickerVol = ws.Cells(x, 9).Value
ws.Range("P4").Value = GreatTotalVolume
ws.Range("O4").Value = TickerVol

End If

Next x

Next ws

End Sub
