# Stock-Analysis

## Overview of Project
Steve wants to find what is the rate of return for specific stocks from a year to year basis for his parents. In this situation we have the data for 2017 and 2018.

##Result
in 2017, The stocks Steve chose for his parents, as shown in the image below, had a good rate of return. All the stocks, besides TERP (-5.4%), provided a solid return, with DQ providing more than 199.4% year to year return.

![2017 Stock Rate of Return](https://user-images.githubusercontent.com/111706055/189501804-ad8a62d1-57bb-4e59-b133-9117eb786037.png)

in 2018 however, the stock market took a turn for the worst. Most of the stocks Steve chose for his parents were negative, with DQ falling by more then 62.6%. However, RUN and ENPH still provided a solid rate of return of 84%, and 81.9& respectively. 

![2018 Stock Rate of Return](https://user-images.githubusercontent.com/111706055/189501807-7ef8db99-ab3b-4d13-bdfe-0337ca29bde9.png)

The analysis can be found in the following link https://github.com/kiwidata/Stock-Analysis/blob/main/VBA_Challenge.xlsm.xlsm

Furthermore please find enclosed the code used for the rate of return of the stocks steve chose for his parents

Sub AllStocksAnalysisRefactored()
    
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("PIE").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
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
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    
    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerEndingPrices(12) As Single
    Dim tickerStartingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For j = 0 To 11
    tickerVolumes(j) = 0
    tickerEndingPrices(j) = 0
    tickerStartingPrices(j) = 0
    Next j
    
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            
        'End If
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        'If  Then
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
    
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If

            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1) <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
            
        'End If
        End If

    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("PIE").Activate
    
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
    
    'Formatting
    Worksheets("PIE").Activate
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

## Summary
In general refactoring makes the code more efficient, and allow a better understanding for future reader to read since its usually a polish of the original code. It also usually allow the computer to use less memory and shows faster results. A disavantage is that by restructuring the code we might create some new error that we did not have before.

For this VBA script the biggest advantage of this restructing was at which speed the result was given. Please find below the comparison between the original and the refactor code.

![original code speed of result](https://user-images.githubusercontent.com/111706055/189502732-94dfc373-c504-4c77-b9e5-3b3d7c1753b3.png)

![refactoring code speed of result](https://user-images.githubusercontent.com/111706055/189502738-5ef04dd8-22fa-41ce-a8e3-fd043457df9d.png)

We can clearly see that the refactor code gave a much quicker result.

One disavantage of the refactor code was that it did create multiple bugs that had to be address while creating the code.
