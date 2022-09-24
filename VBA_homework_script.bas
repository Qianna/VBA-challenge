Attribute VB_Name = "Module1"
Sub stock():

    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    ' Set an initial variable for ticker symbol
    Dim Ticker_symbol As String
    
    ' Set an initial value for totl stock volume
    Dim total_volume As Double
    total_volume = 0
    
    ' Keep track of the location for each ticker in the summary table
    Dim summary_table_row As Integer
    summary_table_row = 2
    
    ' Find last row of data
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Set open price row to 2
    j = 2
    
    ' Loop through all tickers
    For i = 2 To last_row
        
        'check if next ticker is the same as the current ticker, if not
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            ' Set ticker symbol
            Ticker_symbol = Cells(i, 1).Value
            
            'calculate change in price
            Cells(summary_table_row, 10).Value = Cells(i, 6).Value - Cells(j, 3).Value
                
                'conditional formatting
                If Cells(summary_table_row, 10).Value > 0 Then
                    Cells(summary_table_row, 10).Interior.ColorIndex = 4
                
                Else
                    Cells(summary_table_row, 10).Interior.ColorIndex = 3
                
                End If
    
            ' calculate percent change, format to %, then write to summary table
            Cells(summary_table_row, 11).Value = Format((Cells(summary_table_row, 10).Value = Cells(i, 6).Value - Cells(j, 3).Value) / Cells(j, 3).Value, "Percent")
            
            ' add to the total volume
            total_volume = total_volume + Cells(i, 7).Value
            
            ' print ticker symbol in the summary table
            Range("I" & summary_table_row).Value = Ticker_symbol
            
            ' print the total volume to the summary table
            Range("L" & summary_table_row).Value = total_volume
            
            'calcu
            
            'add one to the summary_table_row
            summary_table_row = summary_table_row + 1
            
            'set next open price row
            j = i + 1
            
            ' reset total stock volume
            total_volume = 0
        
        Else
            ' Add to total volume
            total_volume = total_volume + Cells(i, 7).Value
        
        End If
        
    Next i
               
            
End Sub
