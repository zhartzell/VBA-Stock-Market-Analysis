# VBA-Stock-Market-Analysis
## Project Overview

### Purpose
The purpose of the current analysis was to help our clients gain a more clear understanding of how different stocks have fluctuated over the a two year span with the goal helping them make informed investment decisions. From large amounts of stock market data, the current code reduces the output to something easily readable. More specifically, it condenses the data down to the yearly volume and yearly percentage return for each stock ticker in both 2017 and 2018. 

### Data
The original data consisted of information on 3012 different stocks in both 2017 and 2017. This information included: start date, starting price, highest price within the year, lowest price within the year, ending price, and trading volume. 

## Results

### Refactoring Process
Starting with the original VBA code, our team meticulously inserted new code that would allow our program to run more efficiently as well as producing the desired information. Using comments to identify what the new lines are doing, our code is easy to read and interpret. Below is an image of the final code we used for this analysis: 

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
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
    
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickersIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If
            

            '3d Increase the tickerIndex.
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
            'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
        
    Next i
    
    'Formatting cells in output sheet
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
    
### Output
The output of this new code in a new sheet in Excel that displays each Stock Ticker type and how it performed and was traded over either 2017 or 2018. When running this code, our client is able to select from either year and easily compare the results. Below are images of the output sheet for both 2017 and 2018. 

<img width="478" alt="Output_Sheet_2017" src="https://user-images.githubusercontent.com/89808050/134827537-921e9e48-5357-42a6-8155-4d0e9095c132.png">
<img width="479" alt="Output_Sheet_2018" src="https://user-images.githubusercontent.com/89808050/134827546-e425b944-e5d1-4cfc-82fb-237ed22bf442.png">

## Summary
### Advantages of Refactoring Code
The advantages of refactoring code is that it allows us to run programming operations in a more effcient and tidy way. Not only do the code changes and code comments make it easier for other coders to understand your work, but refactoring can also lead to less timely analyses. By increasing the efficiency of the code, the computer can operate more quickly and utilize less memory. 
### Disadvantages of Refactoring Code
Some disadvantages of refactoring code is that it can be time consuming and risky. Although the refactored code can often run faster than the original, it can take coders a lot of energy to refactor appropriately. In addition, there can be some risk if the refactored code lacks valid test cases or when the application is very big. In these circumstances, the original VBA script may be less risky and preferable. 

