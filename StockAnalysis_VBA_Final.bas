Attribute VB_Name = "Module1"
Sub StockAnalysis()
    Dim Current As Worksheet
    Dim last_row As Long, summary_table_row As Integer
    Dim i As Long, j As Long
    Dim ticker_name As String
    Dim total_volume As Double
    Dim opening_price As Double, closing_price As Double
    Dim quarterly_change As Double, percent_change As Double
    Dim greatest_increase As Double, greatest_decrease As Double, greatest_volume As Double
    Dim last_row_summary As Long

    For Each Current In ThisWorkbook.Worksheets
        Current.Activate

        ' Keep track of the location for each ticker in the summary table
        summary_table_row = 2
        
        ' Set the header of summary table
        Range("I1") = "Ticker"
        Range("J1") = "Quarterly Change"
        Range("K1") = "Percent Change"
        Range("L1") = "Total Stock Volume"
        Range("O2") = "Greatest % Increase Value"
        Range("O3") = "Greatest % Decrease Value"
        Range("O4") = "Greatest Total Volume"
        Range("P1") = "Ticker"
        Range("Q1") = "Value"
        
        ' Set the last row
        last_row = Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Initialize variables
        total_volume = 0
        opening_price = 0
        closing_price = 0

        ' Loop through all rows to identify unique tickers
        For i = 2 To last_row
            If Cells(i - 1, 1) <> Cells(i, 1) Then
                ' Setting the opening price when the ticker changes
                opening_price = Cells(i, 3).Value
                ticker_name = Cells(i, 1).Value
            Else
                ' Update the total volume for the current ticker
                total_volume = total_volume + Cells(i, 7).Value
            End If

            ' Check if the next row is a different ticker or if it's the last row
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Or i = last_row Then
                ' setting the closing price based on the when the ticker changes
                closing_price = Cells(i, 6).Value
                
                ' Calculate the quarterly change and percent change
                quarterly_change = closing_price - opening_price
                If opening_price <> 0 Then
                    percent_change = (closing_price - opening_price) / opening_price
                Else
                    percent_change = 0
                End If

                ' Print the ticker and its data to the Summary Table
                Range("I" & summary_table_row).Value = ticker_name
                Range("J" & summary_table_row).Value = quarterly_change
                Range("K" & summary_table_row).Value = percent_change
                Columns("K:K").NumberFormat = "0.00%"
                Range("L" & summary_table_row).Value = total_volume

                ' Increment the row for the next ticker
                summary_table_row = summary_table_row + 1
                
                ' Reset total volume for the next ticker
                total_volume = 0
            End If
        Next i
        
        ' After processing all rows, determine greatest increases, decreases, and volumes
        greatest_increase = Cells(2, 11).Value
        greatest_decrease = Cells(2, 11).Value
        greatest_volume = Cells(2, 12).Value
        last_row_summary = Cells(Rows.Count, 10).End(xlUp).Row
        
        For j = 2 To last_row_summary
            ' Format cells based on percent change values
            If Cells(j, 10).Value >= 0 Then
                Cells(j, 10).Interior.ColorIndex = 4
            Else
                Cells(j, 10).Interior.ColorIndex = 3
            End If
            
            ' Determine greatest increase
            If Cells(j, 11).Value > greatest_increase Then
                greatest_increase = Cells(j, 11).Value
                Cells(2, 17).Value = greatest_increase
                Cells(2, 17).NumberFormat = "0.00%"
                Cells(2, 16).Value = Cells(j, 9).Value
            End If
            
            ' Determine greatest decrease
            If Cells(j, 11).Value < greatest_decrease Then
                greatest_decrease = Cells(j, 11).Value
                Cells(3, 17).Value = greatest_decrease
                Cells(3, 17).NumberFormat = "0.00%"
                Cells(3, 16).Value = Cells(j, 9).Value
            End If
            
            ' Determine greatest volume
            If Cells(j, 12).Value > greatest_volume Then
                greatest_volume = Cells(j, 12).Value
                Cells(4, 17).Value = greatest_volume
                Cells(4, 16).Value = Cells(j, 9).Value
            End If
        Next j
        
        Columns("I:Q").AutoFit
    Next Current
End Sub
