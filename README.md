# Stock Market Analysis

## Project Overview

### Background

For this challenge I was asked by a friend Steve to analyze a group of 12 stocks that his parents were looking to invest in. In order to do this I created an interactive workbook using Visual Basic Application (VBA) in 			Excel to provide the return on investment (ROI) and the annual volume of each stock. To help visualize the information the cells which output the data are color cordinated to help quickly see trends in the stock. Buttons 		are provided in the work book so that the code can quickly be run and reset.
		
### Purpose

Steve is happy with the workbook and would like to use it to analyze an larger volume of stocks. The code works well for the current volume of data but runs a bit slow, to handle the larger data set we need to refactor our code so that it runs more efficently. In order to check if our refactored code is producing the correct results in less time, we need to compare our refactored codes run time to the originial codes run time.

## Results

### Refactoring the Code

In order to make the code more efficent we will take a few steps. First, I created a new variable for tickerIndex that was set to equal 0. Next, I made three new arrays: -tickerStartingPrices(12), -tickerEndingPrices(12), and -tickerVolumes(12). These three arrays store data for performance of each stock that we will use in a for loop to initialize the tickerVolumes to 0. Inside the for loop I wrote a script that increases the current tickerVolumes variable and adds the ticker volume for the current ticker. Next I wrote two if-then statements to get the tickerStartingPrices and tickerEndingPrices variables. Lastly I used a for loop to loop through the arrays in order to output the "Ticker," "Total Daily Volume," "Return," in the worksheet.

### Performance of Stock 2017 vs 2018

There are clear differences in the performance of the Green Stocks between the years of 2017 and 2018. While the returns for Green Stocks were good in 2017 (only TERP had negative returns), 2018 is a different story. In 2018 the return on nearly all Green Stocks are negative (only ENPH and RUN returns are positive) as you can see in the charts below. 

![2017 Results](https://github.com/PSWil/stock-analysis/blob/main/Resources/Output_2017.png)
![2018 Results](https://github.com/PSWil/stock-analysis/blob/main/Resources/Output_2018.png)

While the retuns for Green Stocks in 2018 appear much lower than 2017, which they are, it may not be indicitve that these particular stocks are doing bad compared to the rest of the stock market. Through other resources [PBS NewsHour](https://www.pbs.org/newshour/economy/making-sense/6-factors-that-fueled-the-stock-market-dive-in-2018) and [DW](https://www.dw.com/en/2018-the-worst-year-for-stocks-since-financial-crisis/a-46915652) we can see that the entire market took a hit in 2018 and most stocks trended downward. I would suggest that Steves parents look further into the stocks they are interested in and see if there could be some other internal reason for the stock to decrease. If there is not one apparent then Steves parents can have more confidence that these stocks are performing as expected for the market downturn for that time period, and that they could rebound when the market rises again. Now may be a good time to invest in these stocks while the price is deflated. 

### Execution Time

As you can see below, our refactored code works and run much more efficiently. This new code is able to run through all the data in roughly 80% faster than the original code!

![2017 Initial Run Time](https://github.com/PSWil/stock-analysis/blob/main/Resources/Original_Run_2017.png)
![2017 Refactored Run Time](https://github.com/PSWil/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
![2018 Initial Run Time](https://github.com/PSWil/stock-analysis/blob/main/Resources/Original_Run_2018.png)
![2018 Refactored Run Time](https://github.com/PSWil/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)


	
## Summary
		
### Advantages
### Disadvantages 