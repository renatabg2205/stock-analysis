# stock-analysis

## Overview of Project
The purpose of this project is to analyze different green energy stocks by comparing total daily volume (total number of shares traded throughout the day) and yearly return (percentage difference in price from the beginning of the year to the end of the year). The outcome of this analysis will help a family to compare their current stocks' performance with other green energy stocks so they can decide whether to diversify their funds or not.

## Results
The analysis was made using available data for twelve green energy stocks traded in 2017 and 2018. The daily volume for each one of the stocks was summed up to get the total daily volume per year. In order to get the yearly return, the starting price (the stock price on the first day of the year) and the ending price (the stock price on the last day of the year) for each stock were summed up and then the ending price was divided by the starting price and subtracted by 1 to get the percentage of variation between these prices. By having this information, we can easily compare the yearly performance of each green energy stock.

### Original script
The elapsed run time for the first version of this code was 0.4589844 seconds when running it for 2017 and 0.9316406 for 2018.
<br />
![Elapsed run time for the original script (year 2017)](./Resources/VBA_Challenge_2017_before_refactoring.PNG)

![Elapsed run time for the original script (year 2018)](./Resources/VBA_Challenge_2018_before_refactoring.PNG)

[Original script](green_stocks.xlsm)

### Refactored script
The original script ran through all the cells, summing up the values for each stock and printing the results for each ticker. The refactored code has arrays that store this information for each ticker and outputs the data by running through these arrays and printing the results to the spreadsheet. After refactoring the code, the elapsed run time is now 0.1875 seconds.
<br />
![Elapsed run time for the refactored script (year 2017)](./Resources/VBA_Challenge_2017.PNG)

![Elapsed run time for the refactored script (year 2018)](./Resources/VBA_Challenge_2018.PNG)

[Refactored script](VBA_Challenge.xlsm)

## Summary
The advantages of refactoring a code are faster performance and a code that is easier to read and to be maintained (not only by the original developer but by anyone else that joins the project). In order to refactor the code, additional hours are needed which also means additional costs to a project. Also additional tests must be made to make sure the code is working properly. 

By refactoring the original script ([green_stocks.xlsm](green_stocks.xlsm)) and creating the arrays to store daily volume, starting and ending price, the run time is faster and the code has comments making it easier to read and to maintain ([VBA_Challenge.xlsm](VBA_Challenge.xlsm)).
