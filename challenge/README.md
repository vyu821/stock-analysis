# Stock Analysis with VBA

## Overview of Project

### To create a macro to efficiently analyze many stocks at once. Macro will give total daily volume and return percentage based on stock ticker. Refactoring the code will hopefully run faster.

## Results

### Original Code Performance Times
![2017_Time_Results](https://github.com/vyu821/stock-analysis/blob/master/challenge/resources/2017_Time_Results.png)
![2018_Time_Results](https://github.com/vyu821/stock-analysis/blob/master/challenge/resources/2018_Time_Results.png)

### Refactored Code Performance Times
![VBA_Challenge_2017](https://github.com/vyu821/stock-analysis/blob/master/challenge/resources/VBA_Challenge_2017.png)
![VBA_Challenge_2018](https://github.com/vyu821/stock-analysis/blob/master/challenge/resources/VBA_Challenge_2018.png)

### Original Code
```
'first loop goes through tickers array
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        
        'second loop goes through all the rows in yearValue sheet
        For j = 2 To rowEnd
            Worksheets(yearValue).Activate
...
        Next j

    Next i
```

### Refactored Code
```
    '6a) Initialize ticker volumes to zero
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
               
    '6b) loop over all the rows
    For i = 2 To RowCount
    
        '7a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 9).Value
        
        '7b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 2).Value = tickers(tickerIndex) And Cells(i - 1, 2).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = tickerStartingPrices(tickerIndex) + Cells(i, 7).Value
        End If
        
        '7c) check if the current row is the last row with the selected ticker
        If Cells(i, 2).Value = tickers(tickerIndex) And Cells(i + 1, 2).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = tickerEndingPrices(tickerIndex) + Cells(i, 7).Value
            tickerIndex = tickerIndex + 1
        End If
    
    Next i
```

#### Looking at the execution times of the original code and the refactored code, you can see that the refactored code performed considerably faster. The times of 2017 compared to 2018 are similar, most likely due to the fact that both datasets have the same amount of rows. Having the same amount of rows is significant because the `for` loop used goes through the rows.

## Summary

#### One advantage of refactoring code is to increase efficiency. This can be done by taking fewer steps and using less memory. Another advantage is to improve the logic of the code, thus making it easier to read for future users. 

#### The advantage of the refactored code would be the significant decrease in performance time. This is attributed to the absence of the nested `for` loops used in the original code. Looking at the original code, you can see that it goes through all of the dataset's rows 12 times; compared to the refactored code's `for` loop, which only goes through all the dataset's rows once.
