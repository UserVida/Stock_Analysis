# Analyzing Stock on VBA

## Overview of Project

### Purpose

The purpose of this project was to analyze stock performance from a given year using automated code written in Visual Basic Application (VBA). The data on the stocks, presented in two sheets, was accessed through Microsoft Excel. The data in this analysis was performed on green energy stocks for the years 2017 and 2018. The code, known as a macro, was then refactored to support stock analysis on the entire stock market for the past few years. 

## Results 

### Analysis of Code

In order to evaluate the performance of each stock, the analysis required getting the total daily volume and yearly return for each stock from the years 2017 and 2018. The stocks are represented in this data through their tickers. The macro included an input function to input the year in which the analysis would be performed.

The amount for the total daily volume was retrieved by using a code to create a for loop that initialized the variable tickerVolumes to zero and then iterated it through the row of tickers to collect the value of the total volume for each ticker.

Dim tickerVolumes(12) As Long

For i = 0 To 11
        tickerVolumes(i) = 0
Next i

For i = 2 To RowCount
tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

The yearly return for each stock was performed through a calculation. The starting prices and ending prices for each stock was retrieved through an if condition under a for loop as shown in the code below. 

Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single

For i = 0 To 11
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
Next i

For i = 2 To RowCount
If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    
        End If
        
       
            
If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
          End If

    
The calculation to get the yearly return for each stock in the return column of the stock analysis worksheet was: 

Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

The resulting value was then formatted into a percentage value.

Range("C4:C15").NumberFormat = "0.0%"

The original stock analysis code was refactored to improve speed and efficiency to handle larger data sets. The main change to the code included initializing a variable named tickerIndex, which was then used to access the index for the ticker array and the three output arrays: tickerVolume, tickerStartingPrice and tickerEndingPrice. 

The refactored code shaved off a whole second on the execution of the entire macro as seen in the comparison images of the execution times documented by the timer. The time was recorded through a code which utilized a starttime and endtime function and output the runtime in a message box. 

INSERT IMAGES

### Analysis of Results

The total daily volume of the stock refers to the number of stocks that were traded in the market throughout the day. A stock that is often traded will have a price that will more accurately reflect it's value. The yearly return of a stock shows how much a stock grew or shrunk by. 

In 2017, all stocks showed growth except for TERP which shrunk by 7.2%. SEDG saw the most growth by 184.5%. Among all the stocks in the array, there was an average growth of 67.3%. The average of the total daily volume of all the stocks was 263,886,592. DQ had the lowest total daily volume at 35,796,200 while SPWR had the most at 782, 187, 000.

INSERT 2017 IMAGE TABLE

Mean while, in 2018 most stocks saw a shrinkage in their returns except for stocks ENPH and RUN which saw a 81.9% and 84% growth respectfully. DQ had the greatest shrinkage at a percentage of -62.6%. Overall, the stocks had an average shrinkage of -6.2%. The average of the total daily volume among all stocks was 275,503,183.AY had the lowest total daily volume at 83,079,900 while ENPH had the most at 607, 473, 500.

INSTILL 2018 IMAGE TABLE


## Summary

### Refactoring Code: Advantages & Disadvantages

Advantages to refactoring code include faster execution of the code, ability to handle larger sets of data more efficiently and clean and organized code. The refactored code is less likely to require alteration as a result. 

Disadvantages to refactoring code include accidently introducing bugs into the code, especially when working with larger macros. Since both the original and refactored code have the same functionality, editing the code could potentially create confusion for other programmers who view the code. 

### Refactoring the VBA_Challenge Code: Advantages & Disadvantages 

When refactoring the code in the macro, I found the flow in the refactored code appeared clearer and more organized. The runtime for the refactored code was also much faster compared to the original code.

One difficulty I faced refactoring the code was the introduction of the tickerIndex. I confused the tickerIndex and with the counter (i) and I had to physically draw out the flow and function of the code to understand why I was getting an error when I tried to run the code. For example, when trying to get the value for the yearly return, I initially tried to calculate the formula tickerEndingPrice(tickerIndex) - tickerStartingPrice(tickerIndex) - 1. The code would not run at that point so I had to revisit each step to see where I went wrong before I realized that instead of (tickerIndex), I should have been using (i). 
