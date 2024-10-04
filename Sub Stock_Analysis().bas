Sub Stock_Analysis()
    'Create multiple worksheet loop
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        
        ws.Activate
        
        'Find the last row of the table
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).row
        
        'Add headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Quarterly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        'Create variables
        Dim open_price As Double
        Dim close_price As Double
        Dim quarterly_change As Double
        Dim ticker As String
        Dim percent_change As Double
        Dim volume As Double
        Dim row As Long
        Dim column As Integer
        
        volume = 0
        row = 2
        column = 1
        
        'Setting the initial price
        open_price = Cells(2, column + 2).Value
        
        'Loop through all tickers
        For i = 2 To last_row

            If Cells(i + 1, column).Value <> Cells(i, column).Value Then
                'Setting ticker name
                ticker = Cells(i, column).Value
                Cells(row, column + 8).Value = ticker
                
                'Setting close price
                close_price = Cells(i, column + 5).Value

                'Calculate quarterly change
                quarterly_change = close_price - open_price
                Cells(row, column + 9).Value = quarterly_change

                'Calculate percent change
                If open_price <> 0 Then
                    percent_change = quarterly_change / open_price
                Else
                    percent_change = 0
                End If
                Cells(row, column + 10).Value = percent_change
                Cells(row, column + 10).NumberFormat = "0.00%"

                'Calculate total volume per quarter
                volume = volume + Cells(i, column + 6).Value
                Cells(row, column + 11).Value = volume

                'Iterate to the next row
                row = row + 1

                'Reset open price to next ticker
                open_price = Cells(i + 1, column + 2).Value

                'Reset volume for next ticker
                volume = 0
            Else
                volume = volume + Cells(i, column + 6).Value
            End If
        Next i
        
        'Find the last row of the ticker column
        quarterly_change_last_row = ws.Cells(Rows.Count, 9).End(xlUp).row
        
        'Set the cell colors
        For j = 2 To quarterly_change_last_row
            If Cells(j, 10).Value >= 0 Then
                Cells(j, 10).Interior.ColorIndex = 10 ' Green for positive or zero change
            ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3 ' Red for negative change
            End If
        Next j
        
        'Add headers for quarter analysis
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        
        ' Create variables to store tickers for greatest values
        Dim greatest_percent_increase_ticker As String
        Dim greatest_percent_decrease_ticker As String
        Dim greatest_total_volume_ticker As String
        
        'Find the greatest percent increase, percent decrease, and total volume
        Dim percent_changes As Range
        Dim total_volumes As Range
        
        Set percent_changes = Range(Cells(2, 11), Cells(last_row, 11))
        Set total_volumes = Range(Cells(2, 12), Cells(last_row, 12))
        
        ' Calculate max and min for percent changes and total volumes
        greatest_percent_increase = WorksheetFunction.Max(percent_changes)
        greatest_percent_decrease = WorksheetFunction.Min(percent_changes)
        greatest_total_volume = WorksheetFunction.Max(total_volumes)

        ' Find the corresponding tickers for the greatest values
        For i = 2 To last_row
            If Cells(i, 11).Value = greatest_percent_increase Then
                greatest_percent_increase_ticker = Cells(i, 9).Value
            End If
            If Cells(i, 11).Value = greatest_percent_decrease Then
                greatest_percent_decrease_ticker = Cells(i, 9).Value
            End If
            If Cells(i, 12).Value = greatest_total_volume Then
                greatest_total_volume_ticker = Cells(i, 9).Value
            End If
        Next i

        'Place values in cells
        Cells(2, 17).Value = greatest_percent_increase
        Cells(2, 17).NumberFormat = "0.00%"
        Cells(2, 16).Value = greatest_percent_increase_ticker
        Cells(3, 17).Value = greatest_percent_decrease
        Cells(3, 17).NumberFormat = "0.00%"
        Cells(3, 16).Value = greatest_percent_decrease_ticker
        Cells(4, 17).Value = greatest_total_volume
        Cells(4, 16).Value = greatest_total_volume_ticker
        
        ' Autofit to display data (Optional, if needed)
        Columns("I:Q").AutoFit
        
    Next ws
    
MsgBox "A+?"

End Sub


Sub reset_file(): 'Resets all sheets to pre-analysis state
    Dim i As Integer
    
    'Loop to cycle through all workbook sheets and delete columns I through Q - This also resets formating
    For i = 1 To Sheets.Count
        With Sheets(i)
            .Columns("I:Q").Delete
        End With
    Next i
End Sub