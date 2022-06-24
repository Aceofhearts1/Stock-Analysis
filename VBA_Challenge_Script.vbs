Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    'Alex: TIcker doesn't need to be 12 since we have 11 strings but I imagine it does
    'not mess our code up. (11) wouldve given us 12 slots
    Dim tickers(11) As String
    
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
    Worksheets(yearValue).Activate
    
    'Turning the starting point into a variable in case we decide to format the spreadsheet differently one day
    RowStart = 2
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
        
    tickerindex = 0

    '1b) Create three output arrays
        
    Dim tickerVolume(11) As Long
    Dim tickerStartingPrices(11) As Double
    Dim tickerEndingPrices(11) As Double
    Dim StartingDate(11) As Date
    Dim EndingDate(11) As Date
    
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    'Alex: I do this so we don't have to initilize the array the long way by typing each
    'one out and making them equal 0
    'I placed Starting date and ending date in here to switch up how we find the first and last CLose
    'I made the starting date really high because i want to find the lowest date for a startingprice
    'Vise Versa on the ENding DAte. I want the highest date for the ending date
    For i = 0 To 11
    
        tickerVolume(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
        StartingDate(i) = #1/1/3099#
        EndingDate(i) = #1/1/1999#
        
    
    Next i
    
    
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = RowStart To RowCount
    
        '3a) Increase volume for current ticker
        'Doesnt mention to use an if statement but I can't think of how to do this without one
        If Cells(i, 1).Value = tickers(tickerindex) Then
            
            tickerVolume(tickerindex) = tickerVolume(tickerindex) + Cells(i, 8).Value
            
        End If
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'Doing it a little differently. FInding the starting price by finding the smallest date
        If Cells(i, 2).Value < StartingDate(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then
            
            StartingDate(tickerindex) = Cells(i, 2).Value
            tickerStartingPrices(tickerindex) = Cells(i, 6).Value
            
        End If
        
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
         'Me: Once again doing it differently, just to see if I can.
        If Cells(i, 2).Value > EndingDate(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then
            
            EndingDate(tickerindex) = Cells(i, 2).Value
            tickerEndingPrices(tickerindex) = Cells(i, 6).Value
            
        End If
                               

            '3d Increase the tickerIndex.
            'Based on my way, I'm writing the code as if everything isn't in order.
            'So normally I would have the code read as if to run through the entire spreadsheet
            '11 times. 1 for each ticker. But since we're doing this as if it's in order
            'I'll write this line as if it is in order. I also brainstormed writing
            '11 if else statements with for each row, so it only runs through the spreadsheet once.
        If Cells(i + 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then
        
            tickerindex = tickerindex + 1
              
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolume(i)
        Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

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
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
