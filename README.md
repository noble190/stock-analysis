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

2017:
![Refactored run time - 2017](https://github.com/noble190/stock-analysis/blob/main/resources/VBA_Challenge_2017.png)

2018:
![Refactored run time - 2018](https://github.com/noble190/stock-analysis/blob/main/resources/VBA_Challenge_2018.png)

## Summary
