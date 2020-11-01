# Module 2 Challenge

## Project Overview

This project is concerned with using VBA loops and conditional logic to quickly iterate over data in an excel spreadsheet, in order to analayze the performance of a set of stocks.
It also involves a refactoring exercise, meant to demonstrate how code can be written to perform the same task more efficiently.

## Results

### Stock Analysis

VBA allowed us to quickly analyze a set of transactions involving 12 stocks over two years, in order to generate a summary of their performance. 

The year 2017 saw a degree of growth for 11/12 stocks analyzed, with the strongest gains made by DQ and SEDG. Additionally, FSLR and SPWR were the most traded stocks by a considerable margin.

![2017 Stock Performance](https://github.com/noble190/stock-analysis/blob/main/resources/StockPerformance2017.png)

The year 2018 saw decline for 10/12 stocks analyzed. ENPH and RUN were the only stocks that saw growth. 

![2018 Stock Performance](https://github.com/noble190/stock-analysis/blob/main/resources/StockPerformance2018.png)

Looking at the results above, it would appear that ENPH would be the best stock to recommend. It was in the top 3 stocks in 2017, and did not decline with the rest of the stocks in 2018. Naturally, a larger dataset would facilitate a better founded recommendation.

### Refactoring

A part of this exercise has been refactoring a set of code to be more efficient. The original code ran in approximately 0.8 seconds for both 2017 and 2018. After refactoring (details summarized below), our run time was reduced to approximately 0.1 seconds:

![Refactored run time - 2017](https://github.com/noble190/stock-analysis/blob/main/resources/VBA_Challenge_2017.png)

![Refactored run time - 2018](https://github.com/noble190/stock-analysis/blob/main/resources/VBA_Challenge_2018.png)

## Summary

### General Advantages and Disadvantages of Refactoring Code

Refactoring is generally a good practice, as it allows one to iteratively review previously written code to make it more efficient. Speaking from experience, coming back to a piece of code after a length of time has passed often leads me to think of different (and often better) implementations. It's often useful to have help from someone who did not take part in writing the original code, as they may have a fresh perspective untainted by the struggles of the original implementation.

One disadvantage of refactoring could be the time it may take. There's also a risk that refactoring (especially if done at a later time) could introduce new bugs that may not be immediately apparent. To mitigate this risk, it would make sense to have a strategy for testing the code before undertaking extensive refactoring.

### Refactoring Analysis - Module 2 Challenge

The challenge involved a set of transactions for 12 possible stocks. The original code involved a set of nested loops. For each of the 12 stocks, the code would iterate over the full set of transactions and only record values that matched the stock being examined. 

With 12 stocks and 3012 transactions, the original code involved a total of 36,144 iterations.

The refactored code instead only iterates over the set of transactions once, for a grand total of a mere 3012 iterations (approximately 8% of the original amount). This is consistent with a runtime down from 0.8 seconds in the original implementation to 0.1 seconds with refactored code (with some account for overhead caused by extra complexity of maintaining the additional arrays in the refactored code).

In order to achieve this, the implementation makes a key assumption - that the data is sorted by stock ticker, and sorted in the same order as the tickers array in the code. The original code does not have this constraint, and would be able to parse data in any order. 

The user may easily pre-sort the data in excel. Otherwise, a sorting function may need to be implemented in the refactored code to make it more resilient (which would increase the run time).

Working with a smaller dataset, it may make sense to use the original implementation to avoid this constraint.
