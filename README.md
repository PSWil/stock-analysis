# Stock Market Analysis

## Project Overview

### Background

For this challenge I was asked by a friend Steve to analyze a group of 12 stocks that his parents were looking to invest in. In order to do this I created an interactive workbook using Visual Basic Application (VBA) in 			Excel to provide the return on investment (ROI) and the annual volume of each stock. To help visualize the information the cells which output the data are color cordinated to help quickly see trends in the stock. Buttons 		are provided in the work book so that the code can quickly be run and reset.
		
### Purpose

Steve is happy with the workbook and would like to use it to analyze an larger volume of stocks. The code works well for the current volume of data but runs a bit slow, to handle the larger data set we need to refactor our code so that it runs more efficently. In order to check if our refactored code is producing the correct results in less time, we need to compare our refactored codes run time to the originial codes run time.

## Results

### Refactoring the Code

(To make my code more efficient, I created 3 new arrays: -tickerVolumes(12) to hold volume -tickerStartingPrices(12) to hold starting price -tickerEndingPrices(12) to hold ending price

The above 3 arrays store performance data for each stock when a for loop runs analysis on them. The tickers array that I created in the original establishes a ticker symbol that can be called on for each stock.

Matching the 3 performance arrays with the ticker array is done by using a variable called the tickerIndex.

Now that I have created these arrays, I can use Nested For Loops and variables to loop through the data and complete the analysis.)



### Performance of Stock 2017 vs 2018

![2017 Results](https://github.com/PSWil/stock-analysis/blob/main/Resources/Output_2017.png)
![2018 Results](https://github.com/PSWil/stock-analysis/blob/main/Resources/Output_2018.png)

### Execution Time



	
## Summary
		
### Advantages
### Disadvantages 