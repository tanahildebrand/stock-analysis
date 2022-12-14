# Excel VBA Stock Analysis
## Overview of Project
The purpose of this project was to refactor Excel VBA code to determine if performance was improved. We were provided stock performance data for 2017 and 2018 for 12 stocks, which included trading volumes and opening/closing price per day. We created a script to calculate the total trading volume and annual return per stock based on the year entered. The initial code resulted in a run time between 1.27-1.31 seconds.

## Results
The refactored code makes use of the array function, which allowed me to write the program using only one variable, representing each stock index (0-11).

Below is the refactored code:

```
1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    For I = 0 To 11
        tickerVolumes(I) = 0
        tickerStartingPrices(I) = 0
        tickerEndingPrices(I) = 0
    Next I
        
    '2b) Loop over all the rows in the spreadsheet.
    For I = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(I, 8).Value

        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(I, 1).Value = tickers(tickerIndex) And Cells(I - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(I, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
         If Cells(I, 1).Value = tickers(tickerIndex) And Cells(I + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(I, 6).Value
        End If

        '3d) Increase the tickerIndex.
        If Cells(I, 1).Value = tickers(tickerIndex) And Cells(I + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If
    
    Next I
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For I = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + I, 1).Value = tickers(I)
        Cells(4 + I, 2).Value = tickerVolumes(I)
        Cells(4 + I, 3).Value = tickerEndingPrices(I) / tickerStartingPrices(I) - 1
        
    Next I
```

The refactored code reduced the run time to approximately 0.22 seconds, over a second faster that the original code. Below are the run times for the refactored code:

### 2017 Run Time
![2017 Run Time](/Resources/VBA_Challenge_2017.png)

### 2018 Run Time
![2018 Run Time](/Resources/VBA_Challenge_2018.png)

## Summary
### Advantages and Disadvantages of Refactoring Code
Refactoring code is a critical step in the development process. The advantages of refactoring can include improved processing times and detection of bugs. When code is simplified, organized and properly commented, it allows for better understanding and reduces the risk that future changes are improperly done. Refactoring code does have disadvantages. It takes time to refactor, which can lead to project delays and additional cost if not done appropriately.
### Advantages and Disadvantages of the Original and Refactored VBA Script
A disadvantage of the original script is the user experience. Since the results are being entered before the next stock data is calculated, the experience is diminished as the screen appears to flicker during the process. This is greatly minimized in the refactored code with the use of arrays. In addition to the improved run time of the refactored code, the refactored code combines the formatting of the code into one subroutine, making the macro a one-click process. Both VBA Scripts share a couple of disadvantages. The scripts will only run for stocks already present. If a new stock ticker were added to the data, the logic would need to be updated. In addition, if the yearly data is sorted in another method, the logic to calculate the return is no longer accurate as it relies on the current sort method. If the code could be refactored further to eliminate the reliance of the sort, this would reduce the risk that end user changes (i.e. resorting) would break the logic.