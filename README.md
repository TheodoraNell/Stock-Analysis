# VBA Stock Analysis

## Overview
Performing analysis with VBA on daily stock market data for 2017 and 2018

### Purpose
The purpose of this analysis is to compare the overall performance of array of stocks in the green power sector for each year. A refactored code is then compared to the original in regard to function, return time, and application to a larger dataset with more stocks.


## Method
The original dataset contains spreadsheets for each year (2017,2018) with data for the same 12 stock tickers. Each sheet contains a row for almost every day of the year for every stock ticker. Therefore, the goal is to create a VBA macro to loop through all the rows and pull the relevent data into a new sheet ("All Stocks Analysis") showing the total volume and the overall return for each stock, each year.

An initial step is to initialize an array to represpent the 12 stock ticker names. This array is used to assign the data in each row to the correct ticker and return it to the correct column in the analysis sheet. 

> `Dim Tickers(11) As String`

The array is then defined by each ticker before activating the sheet for the selected year and looping through the rows. At this point there are two different methods used to loop through the rows of data and extract the data necessary for the analysis. 

### Method 1 - Using a nested loop 
In order to assign the correct values, the Tickers array is assigned to a variable and then every ticker in the array is checked against every row in the sheet for the selected year. The relevant data for calculating the Total Volume is present in every row so the loop checks for the current ticker and then adds the daily volume to the total before returning the result for the entire year. 

    For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0`
       
       Worksheets(yearValue).Activate
       
       For j = 2 To RowCount
       
           'Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If`
           
Bcause the data is arranged chronologically and grouped by ticker, the Starting Price and Ending Price values for each ticker are located only in the first and last rows for that ticker. Two more conditional statments are added into the loop to get these values by checking if each row has the current ticker and then separately checking if the row before and the row after have a different ticker.

    'Get starting price for current ticker
    If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

       startingPrice = Cells(j, 6).Value

       End If
           
    'Get ending price for current ticker
    If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

       endingPrice = Cells(j, 6).Value

       End If

### Solution 2 - Using multiple arrays

In the refactored code, a Ticker Index is created to assign the correct values to each ticker rather than looping through all the rows in the data for every ticker in the Tickers array. In this method the variables for the output values are declared as arrays corresponding to the ticker index.

    'Create a ticker Index
    tickerIndex = 0
        
    'Create three output arrays
    Dim tickerVolumes(11) As Long
    Dim tickerStartingPrices(11) As Single
    Dim tickerEndingPrices(11) As Single
    
A single loop is used to iterate though all the rows in the data. First, to cumulatively add up the Total Volume for each ticker without the need for a conditional statment.

    'Create a for loop to initialize the tickerVolumes to zero
    For i = 0 To 11
    
        tickerVolumes(i) = 0
        
    'Loop over all the rows in the spreadsheet
    For i = 2 To RowCount
    
        'Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
The loop then uses conditional statments to check if each row is either the first or the last row of the current Ticker before returning the Starting and Endidng Prices for that Ticker and increasing the index to the next one.

    'Check if the current row is the first row with the selected tickerIndex
    If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
    
       'Get the starting price for the current tickerIndex
       tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                
    End If
        
    'Check if the current row is the last row with the selected ticker
    If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
    
       'Get the ending pricer for the current tickerIndex
       tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
    
        
       'Increase the tickerIndex to the next ticker
       tickerIndex = tickerIndex + 1
            
            
    End If
    
This refactored method not only simplifies the code, but also decreases the amount of time it takes to run. With an array of only 12 stocks and sheets containing around 3000 rows, this decrease in run time can still be seen significantly. The next section explores the differences in the refactored code and how this decrease in run time is important when analyzing much larger datasets. 

## Results

