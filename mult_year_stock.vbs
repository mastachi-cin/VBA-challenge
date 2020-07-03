Sub mult_year_stock()
    Dim ticker As String
    Dim opening_price As Double
    Dim closing_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim grand_total As LongLong
    Dim first_row_ticker As Boolean
    Dim great_inc_ticker As String
    Dim great_inc_val As Double
    Dim great_dec_ticker As String
    Dim great_dec_val As Double
    Dim great_vol_ticker As String
    Dim great_vol_val As LongLong
   
    ' Loop through all sheets
    For Each ws In Worksheets
   
        'Activate current sheet
        ws.Activate
       
        ' Determine the Last Row within the sheet
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        result_table_row = 2
       
        'Add headers for Results Table
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volumen"
       
        'First row by ticker
        first_row_ticker = True
       
        'Loop through all rows within the sheet
        For row_no = 2 To last_row
            'Set opening price from first row by ticker
            If first_row_ticker = True Then
                opening_price = Range("C" & row_no).Value
                first_row_ticker = False
            End If
           
            If Cells(row_no, 1).Value <> Cells(row_no + 1, 1).Value Then
                'Add row to Results Table
               
                'Print Ticker
                ticker = Range("A" & row_no).Value
                Range("I" & result_table_row).Value = ticker
               
                'Print Yearly Change
                closing_price = Range("F" & row_no).Value
                yearly_change = closing_price - opening_price
                Range("J" & result_table_row).Value = yearly_change
               
                ' Highlight negative change in red
                If yearly_change < 0 Then
                    'Range("J" & result_table_row).Font.ColorIndex = 1
                    Range("J" & result_table_row).Interior.ColorIndex = 3
                End If
               
                ' Highlight positive change in green
                If yearly_change > 0 Then
                    'Range("J" & result_table_row).Font.ColorIndex = 1
                    Range("J" & result_table_row).Interior.ColorIndex = 4
                End If
               
                'Print Percent Change
                'Avoid division by zero when opening price = 0
                If CStr(opening_price) = "0" Then
                    If CStr(closing_price) = "0" Then
                        percent_change = 0
                    Else
                        percent_change = 1
                    End If
                Else
                    percent_change = yearly_change / opening_price
                End If
                Range("K" & result_table_row).Value = percent_change
                Range("K" & result_table_row).NumberFormat = "0.00%"
               
                'Print Total Stock Volumen
                grand_total = total_stock + Range("G" & row_no)
                Range("L" & result_table_row).Value = grand_total
               
                'Set Greatest % increase
                If percent_change > 0 And percent_change > great_inc_val Then
                    great_inc_val = percent_change
                    great_inc_ticker = ticker
                End If
               
                'Set Greatest % decrease
                If percent_change < 0 And percent_change < great_dec_val Then
                    great_dec_val = percent_change
                    great_dec_ticker = ticker
                End If
               
                'Set Greatest total volumen
                If grand_total > great_vol_val Then
                    great_vol_val = grand_total
                    great_vol_ticker = ticker
                End If
               
                'Add one to results table row
                result_table_row = result_table_row + 1
               
                'Reset total stock volumen
                total_stock = 0
               
                'Reset first row by ticker
                first_row_ticker = True
            Else
                total_stock = total_stock + Range("G" & row_no)
            End If
           
        Next row_no
       
        'Challenge
        'Print headers
        Range("O1").Value = "Ticker"
        Range("P1").Value = "Value"
       
        'Print Greatest % increase
        Range("N2").Value = "Greatest % Increase"
        Range("O2").Value = great_inc_ticker
        Range("P2").Value = great_inc_val
        Range("P2").NumberFormat = "0.00%"
       
        'Print Greatest % decrease
        Range("N3").Value = "Greatest % Decrease"
        Range("O3").Value = great_dec_ticker
        Range("P3").Value = great_dec_val
        Range("P3").NumberFormat = "0.00%"
       
        'Print Greatest total volumen
        Range("N4").Value = "Greatest Total Volume"
        Range("O4").Value = great_vol_ticker
        Range("P4").Value = great_vol_val
       
        'Reset greatest val
        great_inc_val = 0
        great_dec_val = 0
        great_vol_val = 0
   
    Next ws

End Sub