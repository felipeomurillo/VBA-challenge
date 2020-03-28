Sub VBNStocks()

' --------------------
' Author: Felipe Murillo
' Date: March 25, 2020
'
' Title: VBA Stocks
' Description:
' This function is used to analyze stock market data to output a share summary.
' Summary includes: yearly change, percent change & total stock volume
'
' --------------------

'Declare Variables
Dim current As Worksheet
Dim last_col As Long
Dim last_row As Long
Dim diff_tick_counter As Long
Dim early_date As Long
Dim open_price As Single
Dim close_price As Single
Dim ticker_loc As Integer
Dim share_vol As Single

'Initialize variables
early_date = 99991231 'Dec 31, 9999, used to determine opening data (makes initial entry the earliest date)
share_vol = 0

'Cycle thru each worksheet
For Each current In Worksheets
' Loop through all of the worksheets in the active workbook
' and copy their names
    current.Activate

    'Initialize values within worksheet
    diff_tick_cntr = 0
    
    'Determine last_row and last_column of current worksheet
    last_col = Cells(1, Columns.Count).End(xlToLeft).Column
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Create new output table headers (if it hasn't been done already)
    If Cells(1, last_col - 3).Value <> "Ticker" And Cells(1, last_col).Value <> "Value" Then
        ticker_loc = last_col + 2
        Cells(1, ticker_loc).Value = "Ticker"
        Cells(1, last_col + 3).Value = "Yearly Change"
        Cells(1, last_col + 4).Value = "Percent Change"
        Cells(1, last_col + 5).Value = "Total Stock Volume"
    Else
        ticker_loc = 9  'ninth-column, assuming fixed input format
    End If
    
    ' Start of extraction and manipulation routine
    For i = 2 To last_row
        'If adjacent cells are in the same ticker
        If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
            'Sum share volume
            share_vol = share_vol + Cells(i, 7)
            
            'Pull opening date information
            If Cells(i, 2).Value < early_date Then
                early_date = Cells(i, 2).Value
                open_price = Cells(i, 3).Value
            End If
            
        Else
            diff_tick_cntr = diff_tick_cntr + 1
            
            'Add last entry of share volume
            share_vol = share_vol + Cells(i, 7)
            
            'Pull absolute last closing data
            If Cells(i, 2).Value > Cells(i - 1, 2).Value Then
                late_date = Cells(i, 2).Value
                close_price = Cells(i, 6).Value
            End If
            
             'Write Tick Name
            Cells(diff_tick_cntr + 1, ticker_loc).Value = Cells(i, 1).Value
            
             'Write Yearly Change and Format
            Cells(diff_tick_cntr + 1, ticker_loc + 1).Value = close_price - open_price
            Cells(diff_tick_cntr + 1, ticker_loc + 1).NumberFormat = "0.00"
            
            If (close_price - open_price) < 0 Then
                 'Red if share loses value
                Cells(diff_tick_cntr + 1, ticker_loc + 1).Interior.ColorIndex = 3
            ElseIf (close_price - open_price) > 0 Then
                'Green if share gains value
                Cells(diff_tick_cntr + 1, ticker_loc + 1).Interior.ColorIndex = 4
            Else
                ' Yellow if it remains unchanged (due to no inputs or due to no value delta
                Cells(diff_tick_cntr + 1, ticker_loc + 1).Interior.ColorIndex = 6
            End If
            
            'Write Percent Change and format appropriately (check for division by zero).
            If open_price <> "0" Then
                Cells(diff_tick_cntr + 1, ticker_loc + 2).Value = (close_price - open_price) / (open_price)
            Else
                Cells(diff_tick_cntr + 1, ticker_loc + 2).Value = 0
            End If
            Cells(diff_tick_cntr + 1, ticker_loc + 2).NumberFormat = "0.00%"
            
            'Write share volume
            Cells(diff_tick_cntr + 1, ticker_loc + 3) = share_vol
            
            ' Reset tick parameters before moving on to the next share analysis
            early_date = 99991231
            share_vol = 0
            
        End If
    Next i 'row


    ' Write Sheet Summary
    
    ' -- Header --
    Cells(1, ticker_loc + 7).Value = "Ticker"
    Cells(1, ticker_loc + 8).Value = "Value"
    Cells(2, ticker_loc + 6).Value = "Greatest % Increase"
    Cells(3, ticker_loc + 6).Value = "Greatest % Decrease"
    Cells(4, ticker_loc + 6).Value = "Greatest Total Volume"
    ' --
     
     'Initialize max and min values to 1st entry
    max_p = Cells(2, ticker_loc + 2).Value
    max_tick = Cells(2, ticker_loc).Value
    min_p = Cells(2, ticker_loc + 2).Value
    min_tick = Cells(2, ticker_loc).Value
    max_vol = Cells(2, ticker_loc + 3).Value
    max_vol_tick = Cells(2, ticker_loc).Value
    
    ' Determine new max and min values
     For j = 2 To diff_tick_cntr + 1
        
        If Cells(j, ticker_loc + 2) >= max_p Then
           max_p = Cells(j, ticker_loc + 2).Value
           max_tick = Cells(j, ticker_loc).Value
        End If
        
       If Cells(j, ticker_loc + 2) <= min_p Then
            min_p = Cells(j, ticker_loc + 2).Value
            min_tick = Cells(j, ticker_loc).Value
       End If
            
       If Cells(j, ticker_loc + 3) >= max_vol Then
            max_vol = Cells(j, ticker_loc + 3).Value
            max_vol_tick = Cells(j, ticker_loc).Value
        End If
        
     Next j
    
    ' -- Maximum % --
    Cells(2, ticker_loc + 7).Value = max_tick
    Cells(2, ticker_loc + 8).Value = max_p
    Cells(2, ticker_loc + 8).NumberFormat = "0.00%"
    ' -- Minimum % --
    Cells(3, ticker_loc + 7).Value = min_tick
    Cells(3, ticker_loc + 8).Value = min_p
    Cells(3, ticker_loc + 8).NumberFormat = "0.00%"
    ' -- Max Volume --
    Cells(4, ticker_loc + 7).Value = max_vol_tick
    Cells(4, ticker_loc + 8).Value = max_vol
    
    ' -- Resize Columns
    Cells.Select
    Cells.EntireColumn.AutoFit

Next ' Sheet

End Sub