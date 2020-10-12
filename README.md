# VBA Refactoring Analysis

## Overview

I have a friend named Steve. And while I've already built a button-loaded worksheet for my friend to analyze the performance of green energy stocks during 2018-2019, still he wants more. Steve wants to analyze the performance of all the stocks from 2018-2019. The goal is to refactor the code so that the script runs faster and can handle an expanded dataset. Secondary goal: convince Steve to compensate me for all this work--I mean how close are we? 

### Results

To speed up the code I created a variable and three arrays to store the data for our script to loop through istead of relying on referencing the data sheet itself each time. That code looks like:

That code:

    'variable
    Dim tickerIndex As String
        tickerIndex = 0
    'arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
With the arrays defined, need to create nested loops to comb throught the data sheets and store values in the new array. 
    
        'Loop over all the rows in the spreadsheet
        For i = 2 To RowCount
        'Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        
        'Check if the current row is the first row
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
         tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        'Check if the current row is the last row   
         If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
          tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
          tickerIndex = tickerIndex + 1
            
        End If

Lastly we need to collate the data onto our analysis spreadsheet with another for loop:

        For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
    
       Cells(4 + i, 1).Value = tickers(i)
       Cells(4 + i, 2).Value = tickerVolumes(i)
       Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

Now let's run the code and check the performance speed against the original code.
Performance speed for original code:

![2017](https://github.com/watsonlarry/stock-analysis/blob/main/Resources/Stocks%202017.png)
![2018](https://github.com/watsonlarry/stock-analysis/blob/main/Resources/Stocks_2018.png)

Performance speed for refactored code:

![2017r](https://github.com/watsonlarry/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
![2018r](https://github.com/watsonlarry/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

The refactored code cut -1.20sec from the 2017 analysis time and -1.32sec from the 2018 analysis time. Mission success, the code is working at a faster pace and should be able to handle the larger datasets Steve is imagining. 

### Summary
1. Refactoring code can have many advantages. By simplifying the code you make it easier to reuse, modify, or even return to (if some time passes between you and a project). By making the code more elegant and easier to read, your code becomes more welcoming to others. If you're planning on passing off your code to a partner or giving it to others to use, readability is helpful. There are risks though. Anytime you attempt to modify already working code, you risk breaking the original code. Though as long as you've saved the functional code in a separate location I don't see the risk being that great. 

2. The primary goal was a success--the refactored VBA script runs faster. Thus, one advantage of refactoring is immediately apparent: you can increase the processing speed. With regard to the simplification of the script post-refactoring, I think the positives are less certain. Working through the module I thought that writing the script with the given instructions was clear enough. As I built the multiple new arrays and for-loops for this challenge, I found that the syntax becomes more complex and harder to read. While it seems that the end result was worth the refactoring, I think the alterations made the VBA syntax harder to parse.
