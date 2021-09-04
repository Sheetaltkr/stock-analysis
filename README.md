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

#### In 2017
- The best performing stock is **DQ** with yearly return of **199.4%**
- The worst performing stock is **TERP** with yearly return of **-7.2%**
- The overall stock performance has been great for all stocks except TERP

![Stock Performance 2017](https://github.com/Sheetaltkr/stock-analysis/blob/main/resources/VBA_Challenge_output_2017.png)

#### In 2018 
- The best performing stock is **RUN** with yearly return of **84%**
- The worst performing stock is **DQ** with yearly return of **-62.6%**
- The overall stock performance has been poor for all stocks except ENPH and RUN 

![Stock Performance 2018](https://github.com/Sheetaltkr/stock-analysis/blob/main/resources/VBA_Challenge_output_2018.png)

#### Comparison 2017 vs 2018

- The daily volume has increased for the stocks DQ, ENPH,HASI, RUN,SEDG,TERP and VSLR
- The daily volume decreased for the stocks AY,CSIQ,FSLR,JKS,SPWR
- The stocks which performed well in 2017 have poor returns in 2018

![Stock Performance comparison 2017 vs 2018](https://github.com/Sheetaltkr/stock-analysis/blob/main/resources/VBA_Challenge_Comparison_2018_vs_2017.png)

### 1) Original script and Re-factored script execution time comparison for 2017 and 2018

#### Execution time comparison for 2017


## Pros and Cons of Re-Factoring Code

-Detailed Summary

Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job.


## Pros and Cons of Original and Re-factored script

-Detailed Summary
