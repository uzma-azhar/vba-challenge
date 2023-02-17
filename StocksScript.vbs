Attribute VB_Name = "Module1"
Sub stocks():
    
    For Each ws In Worksheets
    
        ' name new columns
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
    
        ' format to autofit
        ws.Range("A:Q").Columns.AutoFit
        
        ' find last row
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
        
        ' track changes for ticker (column A) and display tickers names in column I
        ' add up volume (column G) to get Total Stock Volume in column L
        Dim tickerName As String
        Dim volumeTotal As Double
        volumeTotal = 0
        
        Dim tickerRows As Integer
        tickerRows = 2
        
        Dim row As Long
        
        Dim openRate As Double
        Dim closeRate As Double
        
        Dim yearlyChange As Double
        Dim percentChange As Double
        
        openRate = ws.Cells(2, 3).Value ' first open rate
        
        For row = 2 To lastRow
            
            If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
                tickerName = ws.Cells(row, 1).Value ' loop through rows to check for ticker change
                volumeTotal = volumeTotal + ws.Cells(row, 7).Value ' calculate volume change
                ws.Cells(tickerRows, 9).Value = tickerName
                ws.Cells(tickerRows, 12).Value = volumeTotal
                volumeTotal = 0
                
                closeRate = ws.Cells(row, 6).Value ' grab close rate before change
                yearlyChange = closeRate - openRate ' calculate yearly change
                ws.Cells(tickerRows, 10).Value = yearlyChange
                
                    If yearlyChange > 0 Then ' conditional formatting to show positive and negative change
                        ws.Cells(tickerRows, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(tickerRows, 10).Interior.ColorIndex = 3
                    End If
                
                percentChange = (closeRate - openRate) / openRate ' calculate percent change
                ws.Cells(tickerRows, 11).Value = percentChange
                ws.Cells(tickerRows, 11).NumberFormat = "0.00%" ' formatting to percentage
                
                openRate = ws.Cells(row + 1, 3).Value ' grab new open rate after change
                
                tickerRows = tickerRows + 1
            Else
                volumeTotal = volumeTotal + ws.Cells(row, 7).Value
            End If
           
        Next row
        
        'find last row for new columns
        lastRow2 = ws.Cells(Rows.Count, 11).End(xlUp).row
        
        ' find greatest increase, greatest decrease, and greatest total volume
        Dim grIncrease As Double
        grIncrease = 0
        Dim grDecrease As Double
        grDecrease = 0
        Dim grTotVol As Double
        grTotVol = 0
        
        Dim percentRow As Integer
        percentRow = 2
        
        For row = 2 To lastRow2
        
            If ws.Cells(row, 11) > grIncrease Then
                grIncrease = ws.Cells(row, 11).Value
                ws.Range("Q2").Value = grIncrease
                ws.Range("Q2").NumberFormat = "0.00%"
            End If
            
            If ws.Cells(row, 11) < grDecrease Then
                grDecrease = ws.Cells(row, 11).Value
                ws.Range("Q3").Value = grDecrease
                ws.Range("Q3").NumberFormat = "0.00%"
            End If
            
            If ws.Cells(row, 12) > grTotVol Then
                grTotVol = ws.Cells(row, 12).Value
                ws.Range("Q4").Value = grTotVol
            End If
        
            percentRow = percentRow + 1
        Next row
    
        ' find ticker match of each value
        Dim match_increase As Double
        match_increase = Application.Match(ws.Range("Q3").Value, ws.Range("K2:K3001"), 0)
        ws.Range("P2").Value = ws.Range("I" & match_increase + 1)
        
        Dim match_decrease As Double
        match_decrease = Application.Match(ws.Range("Q3").Value, ws.Range("K2:K3001"), 0)
        ws.Range("P3").Value = ws.Range("I" & match_decrease + 1)
        
        Dim match_total As Double
        match_total = Application.Match(ws.Range("Q4").Value, ws.Range("L2:L3001"), 0)
        ws.Range("P4").Value = ws.Range("I" & match_total + 1)
    
    Next ws
End Sub

