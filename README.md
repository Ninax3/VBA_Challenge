# VBA_Challenge

## Overview of Project

The purpose of this analysis is to view total volume and return for years 2017 and 2018 for green stock tickers provided. The client requested a performance evaluation of the green stocks provided for year 2017 and 2018 using a macro with a user-friendly button for end-user ease and accuracy. 

## Results

Performance results show 2017 was an overall better performing year than 2018 for the same green stocks as shown with positive returns show in green cells and negative returns show in red cells. See analysis below for visual. 

https://github.com/Ninax3/VBA_Challenge/blob/main/VBA_Challenge_2017_Analysis.pdf
https://github.com/Ninax3/VBA_Challenge/blob/main/VBA_Challenge_2018_Analysis.pdf


The “AllStocksAnalysisRefactored” script was notably 4 times quicker than the original “AllStocksAnalysis” script at about .8 (unrefactored) vs .2 (refactored) seconds. This is due to the refactored code which has more concise steps and loops using the “Index” code.

*Sample of Refactored Code using Index:*
    
    For i = 2 To RowCount
       tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Valu
        
    If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
       tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

## Summary

The advantages of refactoring code is it becomes easier to understand and read. It also becomes less complex and easier to maintain.
The disadvantages of refactoring code is it takes time to edit (refactor) the code that is already written and usable. There is no guarantee of success. 

In relation to the VBA_Challenge scripts, the refactored code is more concise and less wordy. As a result, the script performs better as it takes less time to complete. In addition, the refactored code makes it easy to add more stocks to the index and/or more entries to each index. Another advantage with the refactored code is, one button can be created to include all of the elements the client requested displaying user-friendly data. A disadvantage of the original code is it did not include formatting of cells and had it as a separate subroutine. 

https://github.com/Ninax3/VBA_Challenge/blob/main/VBA_Challenge_2017.png
https://github.com/Ninax3/VBA_Challenge/blob/main/VBA_Challenge_2018.png
