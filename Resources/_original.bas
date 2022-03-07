Attribute VB_Name = "VBA_Module"
Sub AllStocksAnalysis()
    Dim startTime As Single
    Dim endTime  As Single

    

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks 2017"
    Worksheets("All Stocks Analysis").Activate
    
    'Create a header row'
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    

    'create empty array for each variable'
    Dim totalVolume(11)
    Dim startingPrice(11)
    Dim endPrice(11)
    Dim tickers(11) As String
    
    'populate array with each ticker'
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets("2017").Activate
    rowStart = 2
    rowEnd = Range("A1").End(xlDown).Row
    
    'loop though and update each array'
    For i = 0 To 11
        
        ticker = tickers(i)
        totalVolume(i) = 0
        startingPrice(i) = 0
        endPrice(i) = 0
        
        For j = rowStart To rowEnd
            'increase totalVolume'
            If Cells(j, 1).Value = ticker Then
                totalVolume(i) = totalVolume(i) + Cells(j, 8).Value
            
            End If
        
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                'set starting price
                startingPrice(i) = Cells(j, 6).Value
             End If
    
            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                'set end price
                endPrice(i) = Cells(j, 6).Value
            End If
        Next j
    
    Next i
    
    
    Worksheets("All Stocks Analysis").Activate
    'loop and fill excel sheet with data from arrays'
        For i = 0 To 11
            Cells(4 + i, 1).Value = tickers(i)
            Cells(4 + i, 2).Value = totalVolume(i)
            Cells(4 + i, 3).Value = endPrice(i) / startingPrice(i) - 1
        
         Next i
    
    'Formatting'
    Worksheets("All Stocks Analysis").Activate
    
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Font.Size = 14
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    'Convert Return Column to Percentage Number Format'
    Columns("c").NumberFormat = "0.00%"
    Columns("A").AutoFit
    Columns("B").AutoFit
    Columns("C").AutoFit
    
    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            Cells(i, 3).Interior.Color = vbGreen
        Else
            Cells(i, 3).Interior.Color = vbRed
        End If
    
    Next i
    
    endTime = Timer
    MsgBox "This un-refactored code ran in " & (endTime - startTime) & " seconds for the year 2017"
End Sub


