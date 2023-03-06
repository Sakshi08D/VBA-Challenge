Sub StockData()
'Declare variables
Dim ws As Worksheet
Dim TickerSymbol As String
Dim LastRow As Long
Dim YearlyOpenPrice As Double
Dim YearlyClosePrice As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalVolume As Double
Dim GreatestIncreaseTicker As String
Dim GreatestDecreaseTicker As String
Dim GreatestVolumeTicker As String
Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestVolume As Double
'Loop through each worksheet
For Each ws In ThisWorkbook.Worksheets
    ' Set column headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' Set initial values for variables
    Dim j As Long
    Dim last_row As Long
    Dim ticker As String
    Dim year_open As Double
    Dim total_volume As Double
    Dim i As Long
    
    j = 2
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ticker = ws.Cells(2, 1).Value
    year_open = ws.Cells(2, 3).Value
    total_volume = 0
    greatest_increase = 0
    greatest_decrease = 0
    greatest_volume = 0

    ' Loop through all rows
    For i = 2 To last_row
        ' Check if ticker symbol has changed
        If ws.Cells(i, 1).Value <> ticker Then
            ' Output results for previous ticker
            year_close = ws.Cells(i - 1, 6).Value
            yearly_change = year_close - year_open
            percent_change = yearly_change / year_open * 100
            ws.Range("I" & j).Value = ticker
            ws.Range("J" & j).Value = yearly_change
            ws.Range("K" & j).Value = percent_change
            ws.Range("L" & j).Value = total_volume
            
            ' Check if current ticker has the greatest increase, decrease, or volume
            If percent_change > greatest_increase Then
                greatest_increase = percent_change
                greatest_increase_ticker = ticker
            End If
            
            If percent_change < greatest_decrease Then
                greatest_decrease = percent_change
                greatest_decrease_ticker = ticker
            End If
            
            If total_volume > greatest_volume Then
                greatest_volume = total_volume
                greatest_volume_ticker = ticker
            End If
            
            ' Reset variables for new ticker
            j = j + 1
            ticker = ws.Cells(i, 1).Value
            year_open = ws.Cells(i, 3).Value
            total_volume = 0
        End If
        
        ' Add to total stock volume for current ticker
        total_volume = total_volume + ws.Cells(i, 7).Value
    Next i

    ' Output results for last ticker
    year_close = ws.Cells(last_row, 6).Value
    yearly_change = year_close - year_open
    percent_change = yearly_change / year_open * 100
    ws.Range("I" & j).Value = ticker
    ws.Range("J" & j).Value = yearly_change
    ws.Range("K" & j).Value = percent_change
    ws.Range("L" & j).Value = total_volume

    'Print the headers for the summary table
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"


    ' Output results for greatest increase, decrease, and volume
    ws.Range("P2").Value = greatest_increase
    ws.Range("O2").Value = greatest_increase_ticker
    ws.Range("P3").Value = greatest_decrease
    ws.Range("O3").Value = greatest_decrease_ticker
    ws.Range("P4").Value = greatest_volume
    ws.Range("O4").Value = greatest_volume_ticker
Next ws 'Move to next worksheet
End Sub

