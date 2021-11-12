# Stock Analysis With VBA

## Overview of Project

### Purpose
The purpose of this project was to refactor a code that we already created in Module 2 in order to make the code look cleaner and work more efficiently. We were tasked with collecting certain stock informations from the years 2017 and 2018 in order to help Steve help his parents make smarter investments. 
## Data Provided
The data provided included two charts with 12 different stock values over the years 2017 and 2018. The stock information given was the ticker symbol, the date the stock was issued, the Open and CLose price, the highs and lows of the day, the adjusted close value, and the volume it was traded during that day. The goal of our project was to retrieve the Total Daily Volume throughout the year, the correct ticker symbol, and the total return percentage if you were to buy that stock on the first day of the year and hold it until the end. 
## Results of Analysis
During the refactoring process, we were given a code with subtle comments to help guide us along the way. I noticed we were basically told to do what the Module 2 walk through was, but to make it neat and coincise. Copied below is the instructions, and the code I came up with to help our analysis. 

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
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
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
# Summary 
## Challenges
The biggest challenge I faced along the way was interperting what the instructions were telling me. I had done this before, in a less efficient way, and understood what the code was doing, but to change the way I initially saw it, and try to clean it up was difficult. The VBA microsoft forums helped guide me along the way in part 2a, by telling me to add the '(12)' after the statement. 


##Pros and Cons of Refactoring code
The overall pro is having cleaner code as well as code that is easier to interpert. It is more concise, so those who have to see our project have an easier time looking through it. The con would have to be trying to train your brain to not accept the first code that works. Having to see it in another lense and not getting frustrated when you can't figure it out as well. 

## Refactoring our Stock Analysis
The biggest benefit in refactoring our initial stock analysis was a decrease in macro run time. The initial run time was about .5 seconds, and with our refactored code, the macro ran in less than .1 seconds! Talk about efficiancy. Pictured below are the run times for the 2017 and 2018 stock analysis. 

![This is an image]
![This is an image]
