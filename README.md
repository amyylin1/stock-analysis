# Stock-Analysis

## 1. Overview of Project:  Explain the purpose of this analysis

   The background of the project is to prepare a VBA script to analyze the performace of a dozen stocks on the spreadsheet.  Since computing time is precious and stock market have thousands of stocks.   The purpose of this analysis is to make the original code run more efficently, so that more stocks can be analyzed with less time.  
  

## 2.  Results:  

### Performance bet. 2017 and 2018

### The original code 

The oringal code is less efficient because of its nested loop.  The code contains an outer loop, "i", that will iterate the tickers from 0 to 11. Here, the "totalVolume" is set to zero, so that it will reset for each outer loop.  The code also has an inner loop, "j", that will iterate all the rows from 2 to the last row.  For each of the row, it will undergo three "If ... Then" conditional statements, to check for "the current ticker volumn", "starting price" and "ending price" before it loops over to the next row.  Once the inner loop loops through all the rows.  It will start the next outer loop "i" , and repeat the code over and over again till i = 11.  This nested loop structure take longer computation time to excute. 

    
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

The refactored code is faster because it only have one loop to excute.  First, "tinkerIndex" is set to zero.  Three output arrays are created for the three variables of the 11 tickers.  Loop "i" is created to interate over the tickerVolumes(i), and the intial "tickerVolumes" is reset to zero for each ticker.  Next, the code will loop over all the rows with three "If...Then" conditional statements to check for "tickerVolumes", "tickerStartingPrices", and "tickerEndingPrices" of 


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

   


### Run-time for the orginal code:

<img width="256" alt="Screen Shot 2022-08-02 at 2 12 43 PM" src="https://user-images.githubusercontent.com/108419097/182449588-6a41be9f-a726-4e5a-966e-be58e0d502cf.png">

<img width="254" alt="Screen Shot 2022-08-01 at 7 02 04 PM" src="https://user-images.githubusercontent.com/108419097/182449714-95d94b9b-211d-4db3-bdd7-b57dc89a1697.png">

### Run-time for the refractored code: 

<img width="251" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/108419097/182451381-ed9b69b4-5ea8-4e24-a62a-ea4dc8b139bb.png">

<img width="259" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/108419097/182451393-4877bde6-ad32-470b-b207-c85608ab98bb.png">

## Summary

The original code is more clean and easy to debugged. The variables are not in arrays, and less bookkeeping.  at the every iteration, the result will be display in the result sheet, which means that calculating and output are doing in the same time.  
The refractored code is harder to debugged.  The varaibles are in arrays, it is harder to keep track of all the calculations.  The advantage of calcuating data in arrays makes the output easier. 
