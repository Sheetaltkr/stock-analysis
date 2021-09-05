# VBA challenge - Stock Analysis
## Overview of Project
### Purpose
- To **re-factor** the stock analysis program code to handle **large volume**( > 1k stocks ) of stock market data  while keeping the code functionality intact. 
- Determine if code has better **run time** and optimized **resource** usage. 


#### Background

Steve uses a stock_analysis program which generates Total\_Daily\_Volume and Yearly\_ Return for a set of 12 Stock Tickers for his parents.

Steve wants to expand the scope of his stock_analysis research to include entire stock market over the last few years to do a little more research.  

The existing code works well for a dozen stocks, however it might not work as well for thousands of stocks. And if it does, it may have performance issues.


***NOTE:***
 
	Daily volume is the total number of shares traded throughout the day; it measures how actively a stock is traded.  
	The yearly return is the percentage difference in price from the beginning of the year to the end of the year.

## Results


### 1) Stock Performance analysis and comparison for 2017 and 2018

##### In 2017
- The best performing stock is **DQ** with yearly return of **199.4%**
- The worst performing stock is **TERP** with yearly return of **-7.2%**
- The overall stock performance has been great for all stocks except TERP

![Stock Performance 2017](https://github.com/Sheetaltkr/stock-analysis/blob/main/resources/VBA_Challenge_output_2017.png)

##### In 2018 
- The best performing stock is **RUN** with yearly return of **84%**
- The worst performing stock is **DQ** with yearly return of **-62.6%**
- The overall stock performance has been poor for all stocks except ENPH and RUN 

![Stock Performance 2018](https://github.com/Sheetaltkr/stock-analysis/blob/main/resources/VBA_Challenge_output_2018.png)

##### Comparison 2017 vs 2018

- The daily volume has increased for the stocks DQ, ENPH,HASI, RUN,SEDG,TERP and VSLR
- The daily volume decreased for the stocks AY,CSIQ,FSLR,JKS,SPWR
- The stocks which performed well in 2017 have poor returns in 2018

![Stock Performance comparison 2017 vs 2018](https://github.com/Sheetaltkr/stock-analysis/blob/main/resources/VBA_Challenge_Comparison_2018_vs_2017.png)

### 1) Original script and Re-factored script execution time comparison for 2017 and 2018
##### Original Code
https://github.com/Sheetaltkr/stock-analysis/blob/main/resources/original_code.bas
##### Re-factored Code
https://github.com/Sheetaltkr/stock-analysis/blob/main/resources/Refactored_code.bas

##### Execution time comparison for 2017
Before-->
![VBA_Challenge_2017_before](https://github.com/Sheetaltkr/stock-analysis/blob/main/resources/VBA_Challenge_2017_before.png)
After-->
![VBA_Challenge_2017](https://github.com/Sheetaltkr/stock-analysis/blob/main/resources/VBA_Challenge_2017.png)

##### Execution time comparison for 2018
Before-->
![VBA_Challenge_2018_before](https://github.com/Sheetaltkr/stock-analysis/blob/main/resources/VBA_Challenge_2018_before.png)
After-->
![VBA_Challenge_2018](https://github.com/Sheetaltkr/stock-analysis/blob/main/resources/VBA_Challenge_2018.png)

#

## Summary: In a summary statement, address the following questions.

### Advantages and Disadvantages of Re-Factoring Code


Re-Factoring means modifying the code for better performance without changing the functionality.
##### Advantages
- Possible Lesser code lines
- Possible Faster run time
- Optimum usage of resources like memory, CPU cycles
- Better/Simpler code logic
- Higher Code Re-usability
- Easier to Debug
##### Disadvantages
- Negative impact to functionality if not done correctly
- Keeping track of code dependencies for a complex code might be tedious
- Time consuming
- Requires thorough testing and Peer review


### How do these pros and cons apply to re-factoring the original VBA script?

Code changes made and their pros/cons:

##### Pros

- Using common if condition to check of the current `Cell(i,1).value` is same as the ticker values for all the metric calculation removed the code redundancy.
- Using common if condition to check for last ticker row for assigning the endtime and increasing the ticker reduced lines of code.
- Using Single data type instead of double as in original code reduced 4 bytes of variable storage instead. It is always efficient to use smallest data-type when possible for optimum use of memory.
- Using single for loop to go over all the rows only once as compared to nested for loops which went over all the rows 11 times was much more efficient.
- Using single index variable to track the ticker, the `total_daily_volume, starttime and endtime` for each ticker was better than using 4 different variables to go through 4 different arrays.
- Use of array made it possible to store the output matrices for each Ticker instead of updating these results  in the cells every time while traversing through each ticker. With arrays we updated the output in the cells at once.
- We changed the output format to 0.0% for simpler representation of the output data.

##### Cons

- Refactoring of code required additional time to come up with better for loop logic, making those changes and testing the changes
- The Refactored code assumes the dataset is sorted by Ticker values. If the sort order changes then the code will not work.

