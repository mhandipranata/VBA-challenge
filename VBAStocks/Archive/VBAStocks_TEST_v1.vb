Sub stock_market_analysis()
    
    'Setting the variable and data type
    Dim ticker_symbol As String
    Dim yearly_change As Double
    Dim percentage_change As Double
    Dim total_volume As Long
    
    'Kepping track of the location for each ticker symbol in the summary table
    Dim summary_table_row As Integer
    summary_table_row = 2
    
    'Variables to calculate yearly change
    Dim open_price As Double
    Dim close_price As Double
    
    'Loop through all ticker symbol
    
    For i = 2 To 525
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            'Extract distinct ticker symbol
            ticker_symbol = Cells(i, 1).Value
            
            'Calculating yearly change
            close_price = Cells(i, 6).Value
            open_price = Cells(2, 3).Value
            
            yearly_change = close_price - open_price
            
            'Calculating percentage change
            percentage_change = yearly_change / open_price
               
            MsgBox (yearly_change)
            
        End If
    
    yearly_change = 0
    close_price = 0
    open_price = Cells(i + 1, 3).Value
    
    
    
    Next i
    
End Sub
