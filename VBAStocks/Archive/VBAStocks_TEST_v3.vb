Sub stock_market_analysis()
    
    'Setting the variable and data type
    Dim i As Long 
    Dim ticker_symbol As String
    Dim yearly_change As Double
    Dim percentage_change As Double
    Dim total_volume As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim start As Long
    Dim summary_table_row As Integer
    Dim ws As Worksheet

    'Challenge: Setting the variable and data type
    Dim min_ticker_symbol As String
    Dim max_ticker_symbol As String
    Dim min_percentage_change As Double
    Dim max_percentage_change As Double
    Dim max_total_volume As Double
    Dim max_volume_ticker As String
    
    'Loop through all sheets
    For Each ws In Worksheets
        
        yearly_change = 0
        percentage_change = 0
        total_volume = 0
        start = 2
        summary_table_row = 2

        'Determine the Last Row
        Dim last_row As Long
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Determine the header for the summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'Challenge: Determine the header for the summary table
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

        'Loop through all ticker symbol
        For i = 2 To last_row
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'Extract distinct ticker symbol
                ticker_symbol = ws.Cells(i, 1).Value
            
                'Add to total stock volume
                total_volume = total_volume + ws.Cells(i, 7).Value
            
                If total_volume = 0 Then
                    'the results will be all 0
                    ws.Range("I" & summary_table_row).Value = ws.Cells(i, 1).Value
                    ws.Range("J" & summary_table_row).Value = 0
                    ws.Range("K" & summary_table_row).Value = "%" & 0
                    ws.Range("L" & summary_table_row).Value = 0
                
                Else
                    If ws.Cells(start, 3).Value = 0 Then
                        For find_price = start To i
                            If ws.Cells(find_price, 3).Value <> 0 Then
                                start = find_price
                                Exit For
                            End if
                        Next find_price
                        
                    End If


                    'To calculate yearly change, determine close_price and open_price   
                    open_price = ws.Cells(start, 3).Value
                    close_price = ws.Cells(i, 6).Value
                    yearly_change = close_price - open_price
                    percentage_change = Round((yearly_change / open_price) * 100, 2)

                    'start the open price for the next ticker symbol
                    start = i + 1

                    'Print the ticker symbol in summary table
                    ws.Range("I" & summary_table_row).Value = ticker_symbol
            
                    'Print yearly change in summary table
                    ws.Range("J" & summary_table_row).Value = Round(yearly_change, 2)
                    
                    'Print percentage change in summary table
                    ws.Range("K" & summary_table_row).Value = "%" & percentage_change

                    'Print the total stock volume in summary table
                    ws.Range("L" & summary_table_row).Value = total_volume
            
                    
                
                    'Cell color green if yearly change positive value and red for negative value
                    If yearly_change > 0 Then
                        ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
                
                    ElseIf yearly_change < 0 Then
                        ws.Range("J" & summary_table_row).Interior.ColorIndex = 3

                    Else
                        ws.Range("J" & summary_table_row).Interior.ColorIndex = 0

                    End if

                End If
                
                'Add one to the summary table row
                summary_table_row = summary_table_row + 1
                
                'Reset close price, yearly change
                close_price = 0
                yearly_change = 0
                
                
                'Challenge: 
                If percentage_change > max_percentage_change Then
                    max_percentage_change = percentage_change
                    max_ticker_symbol = ticker_symbol
                
                ElseIf percentage_change < min_percentage_change Then
                    min_percentage_change = percentage_change
                    min_ticker_symbol = ticker_symbol

                End If
                
                If total_volume > max_total_volume Then
                    max_total_volume = total_volume
                    max_volume_ticker = ticker_symbol
                End If

                'Print max and min percentage change and max total stock volume in the summary table
                ws.Range("P2").Value = max_ticker_symbol
                ws.Range("P3").Value = min_ticker_symbol
                ws.Range("P4").Value = max_volume_ticker
                ws.Range("Q2").Value = "%" & max_percentage_change
                ws.Range("Q3").Value = "%" & min_percentage_change
                ws.Range("Q4").Value = max_total_volume

                'Reset percent change, total volume
                percentage_change = 0
                total_volume = 0

            Else
                'Add to the total stock volume
                total_volume = total_volume + ws.Cells(i, 7).Value
            
            End If
            
        Next i
    
    Next ws
    
End Sub

