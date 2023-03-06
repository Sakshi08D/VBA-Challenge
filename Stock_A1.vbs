Sub StockAnalysis()
    ' Declare variables
    Dim ticker As String
    Dim year_open As Double
    Dim year_close As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_volume As Double
    Dim last_row As Long
    Dim i As Long
    Dim j As Integer
    
    ' Set column headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    ' Set initial values for variables
    j = 2
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    ticker = Cells(2, 1).Value
    year_open = Cells(2, 3).Value
    total_volume = 0
    
    ' Loop through all rows
    For i = 2 To last_row
        ' Check if ticker symbol has changed
        If Cells(i, 1).Value <> ticker Then
            ' Output results for previous ticker
            year_close = Cells(i - 1, 6).Value
            yearly_change = year_close - year_open
            percent_change = yearly_change / year_open * 100
            Range("I" & j).Value = ticker
            Range("J" & j).Value = yearly_change
            Range("K" & j).Value = percent_change
            Range("L" & j).Value = total_volume
            
            ' Reset variables for new ticker
            j = j + 1
            ticker = Cells(i, 1).Value
            year_open = Cells(i, 3).Value
            total_volume = 0
        End If
        
        ' Add to total stock volume for current ticker
        total_volume = total_volume + Cells(i, 7).Value
    Next i
    
    ' Output results for last ticker
    year_close = Cells(last_row, 6).Value
    yearly_change = year_close - year_open
    percent_change = yearly_change / year_open * 100
    Range("I" & j).Value = ticker
    Range("J" & j).Value = yearly_change
    Range("K" & j).Value = percent_change
    Range("L" & j).Value = total_volume
End Sub
