# VBA Stock Analysis

## Overview of Project
  
### Purpose
  For this project we wanted to help Steve analyze stock data by writing macros with Excel VBA code. By streamlining the analysis with VBA, we were able to look at the stock data from the years 2017 and 2018 and quickly display the percentage of positive or negative returns over that year. However, once we had completed that goal, our challenge was to refactor the code to run even more efficiently than the results of our previous code shown here:
  
  ![Slow_Code_2017](https://user-images.githubusercontent.com/82347825/116831861-e8cedf00-ab7f-11eb-830a-a8b1bb76a524.png)
  ![Slow_Code_2018](https://user-images.githubusercontent.com/82347825/116831863-ec626600-ab7f-11eb-8d94-bd80ac23cf53.png)
  
---

## Results

### Analysis
  To begin refactoring, we used the beginning part of our previously completed code which had established our active worksheets, header rows, ticker arrays and number of rows to loop:
  
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
  
  
  We were then given a checklist of steps to take in refactoring our code. The checklist and code I used to complete the refactoring are as follows:
  
    '1a) Create a ticker Index
        'Setting equal to zero before looping over rows
        
        tickerIndex = 0
        
    '1b) Create three output arrays
        'Creating arrays for tickerVolumes, tickerStartingPrices and tickerEndingPrices
        
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
        
        '3c) Check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            End If
            
        '3d Increase the tickerIndex if the next row's ticker doesn't match the previous row's ticker
             
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
        End If

    Next i
            
     
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
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

    'Coloring cells based on positive/negative returns
    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    'End timer
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    End Sub

---

## Summary

### Advantages and Disadvantages of Refactoring Code
  There are clear advantages to refactoring VBA code: 
    (1) It makes the code run much more efficiently, which could be immensely helpful for large scale programming and computation.
    (2) It also helps make the code much more organized and easier to read, which can be especially useful when sharing the code with others since it will be easier for them to       understand the steps the code follows to accomplish the end goal.
  
  There are, however, disadvantages that I can see stemming from refactoring: 
    Once you have successfully written code to achieve a certain goal but then begin to tinker it by refactoring, you could potentially begin to create errors and bugs while making sure the variables, arrays, loops, etc., are all correctly linked to each other and functioning properly. 
    
    
### Advantages and Disadvantages of Original and Refactored VBA Script
  The main advantage of having refactored our VBA script was the decrease in the code's run time, cutting it by nearly one fifth: where it first ran at nearly one second, the refactored code ran in less than 0.2 seconds:
  
![VBA_Challenge_2017](https://user-images.githubusercontent.com/82347825/116827889-8bc82e80-ab69-11eb-8f8f-3dc07b8a7729.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/82347825/116827894-8ec31f00-ab69-11eb-8b65-f11ab3fbeabc.png)

  Refactoring would make sense if you were working on an extremely large data set that cutting down on the processing time would be beneficial and/or necessary. However, for a small scale project like this with stock data filling only ~3000 lines, diminishing the program run time probably does not matter much to Steve and refactoring would only cause more of a headache to the coder and also potentially lead to errors in an already working code.
