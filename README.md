# stock-analysis.
## Overview of Project: 
When investing our money into securities, we don’t just randomly pick stocks or random stocks within a growing industry (i.e. clean energy, tech, pharmaceuticals, etc.). Investing into single stocks does not always yield a positive return in short term, and sometimes stocks fail long term. Many factors play a role in picking successful stocks for the moneys. It is important to research and analyze the history of the stock and its performance over time to find trends. In this project, thousands of green energy stocks from 2017 and 2018, are analyzed in Visual Basic for Applications (VBA). This is the programming language that codes for Excel and with the right data can be used to analyze the performance of stocks giving us a visualization of the better performing stocks.

### Purpose:
The purpose of this analysis is to compare the performance of green energy stocks in both the year, 2017 and 2018. During the module, we made a basic code that allowed us to chart positive return in green and negative return in red. A code that allowed the user to distinguish a visual for either 2017 or 2018 was added. As well as a code that gave us an elapsed time it took to run the actual code. It was then our job to see if refactoring our code with more concise logic would in turn allow our VBA script to run more efficiently.


## Results: 
After analyzing stocks in both years, 2017 and 2018. It signifies the importance of looking at multiple years. The more data we can analyze, the better decision we can make on a given subject.

### A Comparison of Stock Performance, 2017 and 2018
Below are the charts that visualize the performance of the same stocks in two different years. The chart includes three columns; “Ticker”, “Total Daily Volume”, and “Return”. For those that need clarification; a ticker is an abbreviation that identifies a specific stock, the total daily volume is the amount of shares trade in a 24 hours, and the return is the performance of the stock, measure in a percent, is either a loss or growth. As you can see, there are Excel cells in the “Return” column that are either green or red. This enhances how we visualize the chart and quickly indicates the winners and losers. In both years, there were only two stocks that yielded a positive return in both years.

![Resources/All_Stocks_2017_And_2018_Charts](/Resources/All_Stocks_2017_And_2018_Charts.png)

Although, it is good to look at performance during a single year to discover trends. It could also be beneficial to make an additional chart that looks at the entire history of the stock. Ex. If you notice the row with the ticker “SEDG” you can identify that although it reported a 7.8% loss in the year 2018, it still however has a net positive of 176.7% which is a better performance than the above stock, ticker “RUN”, which has only a return of 89.5% in the last two years.

### A Comparison of VBA Script and Elapsed Run Time 
The original script was written during module 2 as we learned how to apply conditional formatting and for loops to run this analysis. Afterwards, our assignment was to refactor this code into a more concise format that uses the same logic just more efficiently and simplified for the computer to compute the script in a faster run time. The original script is compared to the refactored script below 

![Resources/VBA_Challenge_Refactored_vs_Orginal](/Resources/VBA_Challenge_Refactored_vs_Orginal.png)

### Original vs. Refactored VBA Script for Elapsed Run Time (2017)

![Resources/All_Stocks_2017_Original_Run_Time](/Resources/All_Stocks_2017_Original_Run_Time.png)
![Resources/All_Stocks_2017_Refactored_Run_Time](/Resources/All_Stocks_2017_Refactored_Run_Time.png)

Improved by 2.6484375 seconds

### Original vs. Refactored VBA Script for Elapsed Run Time (2018)

![Resources/All_Stocks_2018_Original_Run_Time](/Resources/All_Stocks_2018_Original_Run_Time.png)
![Resources/All_Stocks_2018_Refactored_Run_Time](/Resources/All_Stocks_2018_Refactored_Run_Time.png)

Improved by 2.5820315 seconds

## Summary: 
- What are the advantages or disadvantages of refactoring code?
The advantage of refactoring code is that it improves the efficiency of the script and allow the user to get their visualization in the front end quicker. Refactoring code could apply to improvement in front end performance in apps that we as a society use day to day. The only disadvantage or I would say difficulty is spending the time to figure out and come up with a better way to write your code which takes concentration and some focus.

- How do these pros and cons apply to refactoring the original VBA script?
These pros and cons were certainly encountered when refactoring the original VBA script. I encountered many issues with the original script and had to constantly debug and figure out how to get all the code to run together without error. The refactored code is undoubtedly a better script and decreased the amount of errors encountered to nearly zero.
