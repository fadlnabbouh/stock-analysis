# Module 2: stock-analysis using VBA

# Overview of Project
The purpose of this project is to help Steve, a good friend, write a VBA script to analyze stock data from 2017 and 2018 of stocks of companies that invest in green energy. After writing the code, it was then refactored in order to improve its speed and efficiency. 

#Results

##Analysis
For the full analysis and VBA script, visit: [Module 2- Stock Analysis](link)
Aftering running the script for 2017 and 2018, I determined that there was large growth in green stocks in 2017, while 2018 saw a general drop in the returns. This was made evident by the number of green cells in the return column in 2017, indicating positive returns, vs the red cells in 2018 indicate loss. 2 stocks grew in both years, "RUN" and "ENPH", with "ENPH" seeing the most growth.

##Refactored Code
In addition to successfully analyzing the stock data, I also managed to successfully refactor the code, significantly reducing its run time. The initial code between the two scripts were identical: setting up the input box, worksheets, table header, and arrays. Then, I created 3 For loops instead of nesting the loops in order to make the code more efficient. I initialized the tickerVolume to 0 here: 
	'1a) Create a ticker Index
    tickerIndex = 0
    

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    
    Dim tickerStartingPrices(12) As Single
    
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i

After that was done, I created a for loop to generate the tickerVolumes, tickerStartingPrices, and tickerEndingPrices variables for each stock (indicated by the tickerIndex). I did not write out the results right away, however, and created two separate For loops to do that in order to reduce run time. 

For loop to generate variables: 
	''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        'End If
            End If
            
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        'If  Then
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            End If

            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
        'End If
    
    Next i

For loops to write out results in a table: 

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

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i

With this refactored code, I was able to reduce the rurn time of the script from 0.695 seconds to around 0.12 seconds (see screenshots below). 

![Refactored Code Time 2017](link)
![Refactored Code Time 2018](link)

#Summary

##Advantages and Disadvantages
The biggest advantage to regactoring code is increasing its efficiency and reducing its run time. Although we only analyzed 12 stocks in this project, the original VBA script would take too long to run if we were to analyze thousands of stocks. Refactoring allows me to reduce the time it takes to run. In addition, another advantage would be having cleaner code. Clean code allows it to be reader much more efficiently by others and to allow me to easily work with the code when needed (to improve it or to fix it). A disadvantage, however, to refactoring code is that it is time consuming. Refactoring is only worth doing if the advantage of a cleaner, more efficient code out weighs the disadvantage of the time consumption.  

##Applying to Stock-Analysis
In this stock analysis, refactoring the code did make it much cleaner and easier to read. In addition, it significantly reduce the scripts run time and improved its effiency. The disadvantage applies here as well in that it took time to refactor the entire code to make more effient.