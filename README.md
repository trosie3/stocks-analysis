# Green Stocks Analysis

## Overview of Project

### Purpose
The purpose was to analyze a list of compiled Green Stocks for the year 2017 and the year 2018 to help Steven show his parents smarter investment choicer, by finding the return value over the course of a specific year. Also, a secondary goal was to create a streamed lined macro that can be expanded with data from more years and stocks.

## Results
-Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script

The results of 2017 analysis shows several stocks DQ, SEDG, ENPH and FSLR with a returns of over 100%, and suggests that TERP (with negative return), AY and RUN with returns under 10% probably aren’t good investment choices. 

The results of 2018 show only ENPH and RUN with positive returns both around 80%, and all others having negative returns with JKS and DQ showing the biggest loss.

However, the combined data suggests that ENPH and RUN both having positive returns both years are the safest investment. Of the two ENPH appears to be the safest as will RUN increase from 5 to 84, ENHP decreased from 130 to 82 which is a smaller variance and higher overall return in the two years. Returns found by utilizing for loops to go through each stock and dividing the ending price of a specific stock (ticker) by the starting price and subtracting one. 

  <i>Returns results found by utilizing for loops to go through each stock and dividing the ending price of a specific stock (ticker) by the starting price and subtracting one, see code below.</i>

By refactoring the code and creating one continuous loop, seen below images, to process the stock data according to my computer I was able to cut run time of the original skirt by about 0.05 seconds, which would add up dramatically if the data set was increased. 
  
    For i = 0 To 11
       
       tickerIndex = tickers(i)
       tickerVolume = 0
        
    'Loop over all the rows in the spreadsheet.
     Worksheets(yearValue).Activate
        For r = 2 To RowCount

        'Check if the current row is the first row with the selected tickerIndex.
                If Cells(r, 1).Value = tickerIndex And Cells(r - 1, 1).Value <> tickerIndex Then
                'set starting price
                tickerStartingPrices = Cells(r, 6).Value
                End If
        
        'check if the current row is the last row with the selected ticker
        'If the next row ticker doesn't match, increase the tickerIndex
                If Cells(r, 1).Value = tickerIndex And Cells(r + 1, 1).Value <> tickerIndex Then
                'set ending price
                tickerEndingPrices = Cells(r, 6).Value
                End If
                
        'Increase volume for current ticker
                If Cells(r, 1).Value = tickerIndex Then
                tickerVolume = tickerVolume + Cells(r, 8).Value
                End If

            '3d Increase the tickerIndex.
        Next r
    
    'Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    Worksheets("All Stocks Analysis").Activate
    
    Cells(4 + i, 1).Value = tickerIndex
    Cells(4 + i, 2).Value = tickerVolume
    Cells(4 + i, 3).Value = tickerEndingPrices / tickerStartingPrices - 1

    Next i	

##Summary
1.	What are the advantages or disadvantages of refactoring code?
2.	How do these pros and cons apply to refactoring the original VBA script?

Advantages for refactoring code into one or fewer subroutine(s) from many is that it allows for better and easier future use by the creator, and is more user friendly to others either utilizing the macro as an end-user or others who many need to go in and tweak the code. It also allows you to find redundancies and way to speed up the program. By refactoring the code from the original script in this project the major advantage is more data can be added, and really more data is needed to truly determine which of these stocks are the least risky, highest risk, etc. in order for an investor to make a more informed decision. 
