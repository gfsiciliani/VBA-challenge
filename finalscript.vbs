Sub comer()

    For Each ws In Worksheets
                
        Dim maxval As Double
        Dim minval As Double
        maxval = 0
        minval = 0
        maxvol = 0
        summary_row = 2
                                
        'time variables
        Dim year As String
        year = ws.Name
        period_init = year + "0102"
        period_end = year + "1231"
        
        'loop through entirety of column A
        For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            stockvol = stockvol + ws.Cells(i, 7).Value          'add stock volume
            
            'detect change in ticker
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
                ticker = ws.Cells(i, 1).Value                   'define ticker symbol
                ws.Range("I" & summary_row).Value = ticker      'populate ticker symbol
                ws.Range("L" & summary_row).Value = stockvol    'Populate Total stock volume
                
            End If
                        
            'define and populate yearly change, percent change
            If ws.Cells(i, 2).Value = period_init Then
                open_price = ws.Cells(i, 3).Value   'defines open price
            ElseIf ws.Cells(i, 2).Value = period_end Then
                close_price = ws.Cells(i, 6).Value  'defines close price
                ws.Range("J" & summary_row).Value = close_price - open_price 'calulates and populates yearly change
                
                'formatting cell
                If ws.Range("J" & summary_row).Value > 0 Then
                    ws.Range("J" & summary_row).Interior.ColorIndex = 4
                ElseIf ws.Range("J" & summary_row).Value < 0 Then
                    ws.Range("J" & summary_row).Interior.ColorIndex = 3
                End If
                
                'Percent change operations--------------------------------------------------
                pct_change = (close_price / open_price) - 1     'Calculate percent change
                With ws.Range("K" & summary_row)
                .Value = pct_change                             'Populate Percent change
                .Value = FormatPercent(.Value)                  'Format cell as percentage
                End With
                                
                'increment summary row and reset volume totaler
                summary_row = summary_row + 1
                stockvol = 0
                
            End If
                    
        Next i
                        
        'Column naming & formatting
        ws.Range("I1").Value = "Ticker Symbol"
        ws.Range("J1").Value = "Yearly change"
        ws.Range("K1").Value = "Percent change"
        ws.Range("L1").Value = "Total stock volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % increase"
        ws.Range("O3").Value = "Greatest % decrease"
        ws.Range("O4").Value = "Greatest total volume"
        
        'Final summary
        For j = 2 To ws.Cells(Rows.Count, 10).End(xlUp).Row
            
            'get min/max values and tickers
            If ws.Cells(j, 11).Value > maxval Then
                maxval = ws.Cells(j, 11).Value
                maxticker = ws.Cells(j, 9).Value
            ElseIf ws.Cells(j, 11).Value < minval Then
                minval = ws.Cells(j, 11).Value
                minticker = ws.Cells(j, 9).Value
            End If
              
            'get maxval
            If ws.Cells(j, 12).Value > maxvol Then
                maxvol = ws.Cells(j, 12).Value
                maxvolticker = ws.Cells(j, 9).Value
            End If
              
              
            'populate and format greatest increase
            With ws.Range("Q2")
            .Value = maxval
            .Value = FormatPercent(.Value)
            End With
            
            'populate and format greatest decrease
            With ws.Range("Q3")
            .Value = minval
            .Value = FormatPercent(.Value)
            End With
            
            'populate and format greatest volume
            With ws.Range("Q4")
            .Value = maxvol
            .NumberFormat = "$#,##0"
            End With
            
            ws.Range("P2").Value = maxticker
            ws.Range("P3").Value = minticker
            ws.Range("P4").Value = maxvolticker
            
        Next j
        
        'format column width
        ws.Columns("I:Q").AutoFit
        
        
    Next ws
    
End Sub