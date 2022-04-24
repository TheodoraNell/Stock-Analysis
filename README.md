# VBA Stock Analysis

## Overview
Performing analysis with VBA on daily stock market data for 2017 and 2018

### Purpose
The purpose of this analysis is to compare the overall performance of array of stocks in the green power sector for each year. A refactored code is then compared to the original in regard to function, return time, and application to a larger dataset with more stocks.


## Method
The original dataset contains spreadsheets for each year (2017,2018) with data for the same 12 stock tickers. Each sheet contains a row for almost every day of the year for every stock ticker. Therefore, the goal is to create a VBA macro to loop through all the rows and pull the relevent data into a new sheet ("All Stocks Analysis") showing the total volume and the overall return for each stock, each year.

An initial step is to initialize an array to represpent the 12 stock ticker names. This array is used to assign the data in each row to the correct ticker and return it to the correct column in the analysis sheet. 

    `Dim Tickers(11) As String`

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
       
Lastly, the data is output for the current ticker to the new spreadsheet. This is done by closing the inner nested loop and setting the values within the outermost loop that iterates through the Tickers array. 

       Next j
       
       'Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

    Next i


### Method 2 - Using multiple arrays

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
        
     Next i  
     
    'Loop over all the rows in the spreadsheet
    For i = 2 To RowCount
    
        'Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
The loop then uses conditional statments to check if each row is either the first or the last row of the current Ticker before returning the Starting and Ending Prices for that Ticker and increasing the index to the next one.

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
       
    Next i
    
The last step is to output the values to the All Stocks Analysis sheet. In the refactored code, this is done with a final loop through the output arrays as opposed to setting the values within the original code's nested loop. 

    'Loop through the arrays
    For i = 0 To 11
    
        'Activate worksheet to output the values
        Worksheets("All Stocks Analysis").Activate
        
        'Output the Ticker
        Cells(4 + i, 1).Value = tickers(i)
        
        'Output the Total Daily Volume
        Cells(4 + i, 2).Value = tickerVolumes(i)
        
        'Output the Return
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
This refactored method not only simplifies the code, but also decreases the amount of time it takes to run. With an array of only 12 stocks and sheets containing around 3000 rows, this decrease in run time can still be seen significantly. The next section explores the differences in the refactored code and how this decrease in run time is important when analyzing much larger datasets. 

## Results

### Stock Performance 2017 vs 2018
Through this analysis, it can be concluded that most of the 12 stocks perormed much better in 2017 than in 2018, both in terms of Total Volume as well as Return. Moreover, in 2017 most stocks showed a positive return and in 2018 most showed a negative return. This invites a deeper exploration into the performance of the green power industry as a whole during this time and possible factors that could explain this trend. 

![2017_Cropped](https://user-images.githubusercontent.com/99051640/164998740-5756e0c8-bc80-4425-a903-302b79ee9cb5.png)
![2018_Cropped](https://user-images.githubusercontent.com/99051640/164998743-4a8fe64a-8e97-42ac-805c-3da3b4e75250.png)

   There are also two stocks that stand out against this trend as having more consistent and/or positive performance. ENPH saw less of a return in 2018 than 2017, however it was still a fairly high return at 81.9%. It also increased in total volume. Other stocks managed to increase in total daily volume in 2018 but still had negative returns. RUN appears to have the strongest performance of all the stocks from 2017 to 2018. There was an increase in Total Daily Volume and a large jump in Return from 5.5% to almost 84%. 
   The data is limited in that it only accounts for two years and further analsys of data for years after 2018 would be beneficial to substantiate any conclusions regarding which stocks would be good options to purchase. 

### Execution times of the original vs refactored script
![2017_Original_Code](https://user-images.githubusercontent.com/99051640/164998377-44efbd7f-3dee-4ba3-a3e9-f97fcf054d7f.png)
![VBA_Challenge_2017](https://user-images.githubusercontent.com/99051640/164998383-94aa5ae9-22f0-4e35-8627-90713eeb375a.png)

![2018_Original_Code](https://user-images.githubusercontent.com/99051640/164998379-dac9104f-7468-4da3-97d7-58e134c9d2ed.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/99051640/164998384-445427f4-2298-4b85-a1a3-d6fbd9788a3a.png)


disadvantages in general: time consuming, create higher levels of complexity even when reducing the number of lines (introducing variables that are arrays, dictionaries, etc. May not be as easy to follow or adapt for other use. 

specifically: Variables are arrays creating multiple possible arguments for variables which can get confusing when referencing them in the code. Both depend on the data being arranged in a specific way but the refactored makes it harder to refactor again if applied to a dataset that is not chronological and grouped

