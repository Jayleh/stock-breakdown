Sub stockRunAll():
    
    Call stockTotal
    Call stockYearPercentChange

End Sub


Sub stockTotal():
    
    Dim ws As Worksheet

    For Each ws In Worksheets
        ' Label new column headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
                
        ' Declare ticker letter
        Dim ticker As String
        
        ' Instantiate stock volume total
        Dim stock_volume As Double
        stock_volume = 0
        
        ' Keep track of location of each ticker for summary table
        Dim summary_table_row As Long
        summary_table_row = 2

        ' For the loop
        Dim i, lr As Long
        lr = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Loop through all tickers
        For i = 2 To lr

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Set ticker letter
                ticker = ws.Cells(i, 1).Value

                ' Print ticker in summary table
                ws.Range("I" & summary_table_row).Value = ticker
                
                ' Increase ticker total
                stock_volume = stock_volume + ws.Cells(i, 7).Value

                ' Print ticker total in summary table
                ws.Range("L" & summary_table_row).Value = stock_volume

                ' Increment summary table row
                summary_table_row = summary_table_row + 1

                ' Reset stock volume total
                stock_volume = 0
            Else
                ' Default as increase total stock volume
                stock_volume = stock_volume + ws.Cells(i, 7).Value
            End If

        Next i

    Next ws

End Sub


Sub stockYearPercentChange():

    Dim ws As Worksheet

    For Each ws In Worksheets

        ' Declare ticker open and close prices
        Dim open_price, close_price As Double

        ' Declare open and close rows
        Dim open_price_row, close_price_row As Long
        
        ' Instantiate a row count for ticker quantity
        Dim row_count As Long
        row_count = 0

        ' Declare yearly and percent changes
        Dim year_change, percent_change As Double
        
        ' Keep track of location of each ticker for summary table
        Dim summary_table_row As Long
        summary_table_row = 2

        ' For the loop
        Dim j, lr As Long
        lr = ws.Cells(Rows.Count, 1).End(xlUp).Row

        For j = 2 To lr
            If ws.Cells(j + 1, 1).Value <> ws.Cells(j, 1).Value Then
                ' Calculate row count by subtracting from close row index
                row_count = j - row_count

                ' Grab close and open price row
                open_price_row = ws.Range("C" & row_count).Row
                ' close_price_row = ws.Range("F" & j).Row
                                
                ' Grab close and open price value, j is already close_price_row
                open_price = ws.Range("C" & open_price_row).Value
                close_price = ws.Range("F" & j).Value
                
                ' Calculate yearly change value
                year_change = close_price - open_price
                
                ' Print yearly change value
                ws.Range("J" & summary_table_row).Value = year_change
                
                ' On divide by zero error, resume to next
                On Error Resume Next
                If year_change / open_price = True Then
                    ' Print error value as 0
                    ws.Range("K" & summary_table_row).Value = 0
                Else
                    ' Calculate percent change value
                    percent_change = year_change / open_price
                    
                    ' Print percent change value
                    ws.Range("K" & summary_table_row).Value = percent_change
                End If
                
                ' Increment summary table row
                summary_table_row = summary_table_row + 1

                ' Reset row count
                row_count = 0
            Else
                ' Increment row count
                row_count = row_count + 1
            End If
            
        Next j

        ' Autofit columns
        ws.Columns("J").AutoFit
        ws.Columns("K").AutoFit
        ws.Columns("L").AutoFit
        
        ' Format column K to percentage
        ws.Columns("K").EntireColumn.NumberFormat = "0.00%"

        ' Conditionaly format yearly change
        Dim k, summary_table_lr As Long
        summary_table_lr = ws.Cells(Rows.Count, "J").End(xlUp).Row

        For k = 2 To summary_table_lr

            If ws.Range("J" & k).Value > 0 Then
                ws.Range("J" & k).Interior.ColorIndex = 10
            Else
                ws.Range("J" & k).Interior.ColorIndex = 9
            End If

        Next k

    Next ws

End Sub
