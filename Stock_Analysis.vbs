Sub Run_On_All_Worksheets()
    Dim Sheet as Worksheet
    For Each Sheet In Worksheets
        Sheet.Select
        Call Stock_Analysis
    Next
End Sub

Sub Stock_Analysis()
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row

    Dim ticker As String

    Dim volume As LongLong
    volume = 0

    Dim table_row As Integer
    table_row = 2

    Dim ticker_open As Double
    ticker_open = Cells(2, 3).Value

    Dim ticker_close As Double

    'Filling in table names
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

    For i = 2 To lastrow
        'If the next cell down is a different ticker
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            'update variables
            ticker = Cells(i, 1).Value
            ticker_close = Cells(i, 6).Value
            volume = volume + Cells(i, 7).Value

            'print ticker name
            Range("I" & table_row).Value = ticker

            'print yearly change
            Range("J" & table_row).Value = ticker_close - ticker_open
            'color cell red for negative, green for positive
            If Range("J" & table_row).Value < 0 Then
                Range("J" & table_row).Interior.ColorIndex = 3
            ElseIf Range("J" & table_row).Value > 0 Then
                Range("J" & table_row).Interior.ColorIndex = 4
            End If

            'print percent change
            Range("K" & table_row).Value = Range("J" & table_row).Value / ticker_open
            Range("K" & table_row).NumberFormat = "0.00%"

            'print volume
            Range("L" & table_row).Value = volume
            table_row = table_row + 1

            'reset volume
             volume = 0

             'get next open
            ticker_open = Cells(i + 1, 3).Value

        'If the next cell is the same ticker
        Else
            volume = volume + Cells(i, 7).Value
        End If
    Next i

    'Comparing tickers
    summary_rows = Cells(Rows.Count, "I").End(xlUp).Row
    Dim best_percent as Double
    best_percent = Cells(2, 11).Value
    Dim worse_percent as Double
    best_percent = Cells(2, 11).Value
    Dim best_volume as LongLong
    best_volume = Cells(2, 12).Value

    For i = 2 to summary_rows
        'search for highest percent
        If Cells(i, 11).Value > best_percent Then
            best_percent = Cells(i, 11).Value
            best_percent_ticker = Cells(i, 9).Value
        End If
        'search for lowest percent
        If Cells(i, 11).Value < worse_percent Then
            worse_percent = Cells(i, 11).Value
            worse_percent_ticker = Cells(i, 9).Value
        End If
        'search for highest volume
        If Cells(i, 12).Value > best_volume Then
            best_volume = Cells(i, 12).Value
            best_volume_ticker = Cells(i, 9).Value
        End If
    Next i

    'fill in comparison table
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"

    Cells(2, 16).Value = best_percent_ticker
    cells(2, 17).Value = best_percent
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(3, 16).Value = worse_percent_ticker
    Cells(3, 17).Value = worse_percent
    Cells(3, 17).NumberFormat = "0.00%"
    Cells(4, 16).Value = best_volume_ticker
    Cells(4, 17).Value = best_volume
    Cells(4, 17).NumberFormat = "0"

    Columns("I:Q").AutoFit
End Sub
