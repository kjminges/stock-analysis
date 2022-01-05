# Module 2 Assignment - Stock Analysis

## Background
After completing the initial workbook for Steve, he expressed concern with its ability to handle larger databases. While the initial workbook summarized data for 11 stocks, there is worry that the workbook would not run as efficiently for thousands of stocks. With Steve's plans to expand his research to include the entire stock market, we were asked to refactor our code to collect the same data with a more efficient code. This will allow for the model to use fewer steps, less memory, and less time to run.

## Analysis and Results
### Impact of Refactored Code
In the original code, we used nested for loops to go through the entire 2017 or 2018 databases for each of the each of the eleven stocks. This is a labor-intensive process since the code is essentially running through all 3000 lines in Excel, 11 times over (essentially running through 33,000 lines). The refactored code uses arrays and an index variable in order to avoid running through the data multiple times. Under the new system, the code effectively goes through the data and uses the index variable to separate and perform metrics for each of the eleven stocks. An example of the code used to go through the data can be found below:

```
    For j = 2 To RowCount

        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
      
	If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
      
            tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
            
        End If

        If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
                     
            tickerEndingPrices(tickerIndex) = Cells(j, 6).Value

            tickerIndex = tickerIndex + 1
            
        End If

    Next j
```

The result is a significant reduction in total run time. Compared to the original code for both the 2017 and 2018 data, the refactored code ran the analysis in around one eighth the time. Copies of the execution times for both the original and refactored code can be found below:

Initial Execution Time

![Initial Execution Time (2017)](https://github.com/kjminges/stock-analysis/blob/main/Resources/VBA_Challenge_2017_Initial_Code.png)
![Initial Execution Time (2018)](https://github.com/kjminges/stock-analysis/blob/main/Resources/VBA_Challenge_2018_Initial_Code.png)

Refactored Execution Time

![Refactored Execution Time (2017)](https://github.com/kjminges/stock-analysis/blob/main/Resources/VBA_Challenge_2017_Refactor.png)
![Refactored Execution Time (2018)](https://github.com/kjminges/stock-analysis/blob/main/Resources/VBA_Challenge_2018_Refactor.png)

For an overview of the stock performance in 2017 and 2018 based on the two codes, including updated run times for each (note that the code was edited slightly to ensure a more "apples-to-apples" comparison between the original and refactored), please see the exhibits below.

![2017 Stock Outcomes](https://github.com/kjminges/stock-analysis/blob/main/Resources/VBA_Challenge_2017_Exhibit.png "2017 Stock Outcomes")
![2018 Stock Outcomes](https://github.com/kjminges/stock-analysis/blob/main/Resources/VBA_Challenge_2018_Exhibit.png "2018 Stock Outcomes")

## Conclusion
In summary, there are obvious advantages and disadvantages to the original and refactored codes. In general, refactoring code allows for the code to be re-written giving authors a second chance at writing code that is better organized or that runs more efficiently. A possible disadvantage is that while editing the code, you could introduce more bugs that were not part of the original code and could be hard to catch. This would ultimately lead to less efficient code and, if not caught early enough, could cause significant delays in delivery.

For our specific code written for Steve, there are advantages and disadvantages to refactoring. As we discussed above, a major advantage to the refactored code is its ability to more efficiently summarize the data. This allows the program to run more quickly and using less memory and could allow Steve to summarize larger sets of data without adding additonal strain. After reviewing the code, one disadvantage is the ability for a novice to come in and understand the refactored code (as compared to the original code). While the nested for loops in the original code is less efficient, it is easier for someone who is new to VBA to understand what is happening and to follow the logic. This could be important if turnover creates a situation where a new associate is asked to update or de-bug the code with little to no understanding of VBA.
