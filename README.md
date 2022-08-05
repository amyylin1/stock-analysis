# Stock-Analysis

## 1. Overview of Project:  Explain the purpose of this analysis

   The background of the project is to prepare a VBA script to analyze the performance of a dozen stocks on the spreadsheet.  Since computing time is precious and the stock market has thousands of stocks.   The purpose of this analysis is to make the original code run more efficiently so that more stocks can be analyzed with less time.  
 

## 2.  Results:  

### Performance of all stocks 

The year 2017 was a pretty good year for growth.  All of the stocks (except ticker "TERP") had increased returns.  Ticker "DQ" had the most growth (199.4% of return).  However, compared to 2017, 2018 was a bad year for growth.  All of the stocks (except tickers "ENPH" AND "RUN") experienced negative growth.  

<img width="717" alt="2017_vs_2018" src="https://user-images.githubusercontent.com/108419097/183156607-afef1953-ca71-422e-92ac-eefcf5fd5e6f.png">


### The original code 

The original code is less efficient because of its nested loop.  The code contains an outer loop, "i", that will iterate the tickers from 0 to 11. Here, the "totalVolume" is set to zero, so that it will reset for each of the outer loops.  The code also has an inner loop, "j", that will iterate all of the rows.  For each row, it will execute three "If ... Then" conditional statements, to check for "the current ticker volume", "starting price" and "ending price" before it loops over to the next row.  Once the inner loop loops through all the rows.  It will start the next outer loop "i", and repeat the code over and over again until i = 11.  This nested loop structure takes longer to run.

    
          '4) Loop through tickers

          For i = 0 To 11
          ticker = tickers(i)
          totalVolume = 0

          '5) loop through rows in the data

          Worksheets("2018").Activate

          For j = 2 To RowCount
              '5a) Get total volume for current ticker
               If Cells(j, 1).Value = ticker Then

                  totalVolume = totalVolume + Cells(j, 8).Value

              End If
              '5b) get starting price for current ticker
              If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                  startingPrice = Cells(j, 6).Value

              End If

              '5c) get ending price for current ticker
              If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                  endingPrice = Cells(j, 6).Value

              End If
          Next j
          
          '6) Output data for current ticker
          Worksheets("All Stocks Analysis").Activate
          Cells(4 + i, 1).Value = ticker
          Cells(4 + i, 2).Value = totalVolume
          Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

      Next i


### Refactored code

The refactored code is faster because it only has one loop to execute.  First, "tinkerIndex" is set to zero.  Three output arrays are created for the three variables of the 11 tickers.  Loop "i" is created to iterate over the tickerVolumes(i), and the initial "tickerVolumes" is reset to zero for each ticker.  Next, the code will loop over all the rows with three "If...Then" conditional statements to check for "tickerVolumes", "tickerStartingPrices", and "tickerEndingPrices" of the curent tickerIndex before it moves on to the next tickerIndex (tickerIndex + 1). 


    '1a) Create a ticker Index
       tickerIndex = 0     

    '1b) Create three output arrays
    Dim tickerVolumes(11) As Long
    Dim tickerStartingPrices(11) As Single
    Dim tickerEndingPrices(11) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 
    Next
           
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
        
        ticker = tickers(tickerIndex)
        
        '3a) Increase volume for current ticker
        If Cells(i, 1).Value = ticker Then
        
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
      
        If Cells(i - 1, 1).Value <> ticker And Cells(i, 1).Value = ticker Then

            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
            
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
         'If  Then
         
          If Cells(i + 1, 1).Value <> ticker And Cells(i, 1).Value = ticker Then

                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

         
        '3d Increase the tickerIndex.
            
            tickerIndex = tickerIndex + 1
            
        End If
            
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
 
     For i = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
     Next


### Compare execution times of the original script and the refactored script:

The original code for 2017 and 2018 took 0.273 and 0.305 seconds, respectively.
The refactored code for 2017 and 2018 took 0.090 and 0.078 seconds, respectively.  
Compared to the original code, the refactored code is about 3-times faster.  


### Run-time for the orginal code (2017 and 2018):

<img width="256" alt="Screen Shot 2022-08-02 at 2 12 43 PM" src="https://user-images.githubusercontent.com/108419097/182449588-6a41be9f-a726-4e5a-966e-be58e0d502cf.png">

<img width="254" alt="Screen Shot 2022-08-01 at 7 02 04 PM" src="https://user-images.githubusercontent.com/108419097/182449714-95d94b9b-211d-4db3-bdd7-b57dc89a1697.png">


### Run-time for the refactored code (2017 and 2018): 

![VBA_Challenge_2017](https://user-images.githubusercontent.com/108419097/183152917-2d7dee3c-cfb9-4938-a078-78567f2089fd.png)

<img width="259" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/108419097/182451393-4877bde6-ad32-470b-b207-c85608ab98bb.png">


## Summary


The refactored code is more efficient.  All of the variables are in arrays, therefore, the output is easier and the code runs faster.  However, the refactored code is more complex.  It is harder to de-construct and de-bugged.   The array output also makes it difficult to track all of the intermediate calculations.  

In contrast, the original code is slower.   The variables are not in arrays, therefore, calculations and outputs are executed simultaneously.  Hence, the longer execution time.  However, it is a cleaner code and easy to de-bugged.  

