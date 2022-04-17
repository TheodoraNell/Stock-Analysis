# VBA Stock Analysis

## Overview
Performing analysis with VBA on daily stock market data for 2017 and 2018

### Purpose
The purpose of this analysis is to compare the overall performance of array of stocks in the green power sector for each year. A refactored code is then compared to the original in regard to function, return time, and application to a larger dataset with more stocks.


## Method
The original dataset contains spreadsheets for each year (2017,2018) with data for the same 12 stock tickers. Each sheet contains a row for almost every day of the year for every stock ticker. Therefore, the goal is to create a VBA macro to loop through all the rows and pull the relevent data into two columns showing the total volume and the overall return for each stock, each year.

An initial step is to initialize an array to represpent the 12 stock ticker names. This array is used to assign the data in each row to the correct ticker and return it to the correct column in the analysis sheet. 

> `Dim Tickers(11) As String`

The array is then defined by each ticker before activating the sheet for the selected year and looping through the rows. At this point there are two different methods used to loop through the rows of data and extract the data necessary for the analysis. 

### Solution 1 - Using a nested loop
The relevant data for calculating the total volume is present in every row because the values are separate daily totals and not cumulative. Therefore the loop must check every row for the relevent ticker and add the daily volume to the total before returning the result for the entire year.

> `For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0`
       
       Worksheets(yearValue).Activate
       
       For j = 2 To RowCount
       
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If'

### Solution 2 - Using multiple arrays


## Results

![Originall_VBA_Challenge_2017](https://user-images.githubusercontent.com/99051640/163731165-46e2fe90-c149-46c4-836a-0e72ceb2bb04.png)
![VBA_Challenge_2017](https://user-images.githubusercontent.com/99051640/163731177-28011c9d-7871-48ca-b7fb-1dab53aebb58.png)

![Original_VBA_Challenge_2018](https://user-images.githubusercontent.com/99051640/163731167-c5737c77-2601-420d-a421-406621ba32e1.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/99051640/163731180-22776bec-c764-4e4c-8edd-39b2208e9992.png)
