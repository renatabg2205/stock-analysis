# stock-analysis

## Overview of Project
The purpose of this project is to analyze different green energy stocks by comparing total daily volume (total number of shares traded throughout the day) and yearly return (percentage difference in price from the beginning of the year to the end of the year) in 2017 and 2018.

## Results
The analysis was made using available data for twelve green energy stocks traded in 2017 and 2018. The daily volume for each one of the stocks was summed up to get the total daily volume per year. In order to get the yearly return, the starting price (the stock price on the first day of the year) and the ending price (the stock price on the last day of the year) for each stock were summed up and then the ending price was divided by the starting price and subtracted by 1 to get the percentage of variation between these prices. By having this information, we can easily compare the yearly performance of each green energy stock.

The elapsed run time for the first version of this code (green_stocks.xlsm) was 0.4589844 seconds when running it for 2017 and 0.9316406 for 2018.
<br />
![Elapsed run time for the original script (year 2017)](./Resources/VBA_Challenge_2017_before_refactoring.PNG)

![Elapsed run time for the original script (year 2018)](./Resources/VBA_Challenge_2018_before_refactoring.PNG)


After refactoring the code, the elapsed run time is now XXX for the year of 2017 and XXXX for 2018.

![Elapsed run time for the refactored script (year 2017)](./Resources/VBA_Challenge_2017.PNG)

![Elapsed run time for the refactored script (year 2018)](./Resources/VBA_Challenge_2018.PNG)


## Summary
Summary: In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?
