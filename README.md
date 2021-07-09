# Stock Analysis with VBA

## Overview of Project

To analyze the stock performance of green stocks between the years 2017 and 2018, as well as refactoring existing code to improve script performance.

### Purpose

The purpose of the project is to provide an analysis of the stock performance of green stocks in the years of 2017 and 2018 using refactored code to improve the visualization of the data as well as improving the speed and performance of the script.

## Results

### Analysis of Stock Performance Between 2017 and 2018
Looking at the data outputs below, it is clear there is a clear disparity between the results between the years of 2017 and 2018. 2017 was a very good year for green stocks, with all but 1 stocks having a positive returns, and with a majority of stocks having a return of at least 25 percent to as high as 199.4 percent. Even the lowest performing stock TERP only has a negative return of about 7.2 percent. This contrasts drastically with performance of green stocks in 2018, where all but 2 stocks have negative returns. In 2018, the only stocks to have improved from the previous year would be RUN and ironically TERP, although TERP still has a negative return. Unfortunately, more context is needed for the reason of the decline of a majority of green stocks, since comparing the decrease in performance ticker by ticker creates very disparate results. For example, DQ decreased from a positive 199.4 percent return in 2017 to a negative 62.6 percent return, a 262 percent disparity while AY decreased from positive 8.9 percent to negative 7.3, a 16.2 percent disparity. Overall, we can conclude that there was a negative trend for most stocks from 2017 to 2018.

![2017 Data](https://github.com/DavidFGitH/stock-analysis/blob/61365ecf70c11224dbe474c2b9bd4b79ccddcb91/Resources/2017_Results.png) ![2018 Data](https://github.com/DavidFGitH/stock-analysis/blob/61365ecf70c11224dbe474c2b9bd4b79ccddcb91/Resources/2018_Results.png)

### Analysis of the Performance of Refactored Script
Looking at the performance of the refactored script below:

![2017 Ref Runtime](https://github.com/DavidFGitH/stock-analysis/blob/61365ecf70c11224dbe474c2b9bd4b79ccddcb91/Resources/VBA_Challenge_2017.png) ![2018 Ref Runtime](https://github.com/DavidFGitH/stock-analysis/blob/61365ecf70c11224dbe474c2b9bd4b79ccddcb91/Resources/VBA_Challenge_2018.png)

and comparing it with the below performance of the original script:

![2017 Old Runtime](https://github.com/DavidFGitH/stock-analysis/blob/61365ecf70c11224dbe474c2b9bd4b79ccddcb91/Resources/AllStockAnalysis_2017_Runtime.png) ![2018 Old Runtime](https://github.com/DavidFGitH/stock-analysis/blob/61365ecf70c11224dbe474c2b9bd4b79ccddcb91/Resources/AllStockAnalysis_2018_Runtime.png)

We can concluded that the overall performance of the refactored script is better than the previous one, even if the refactored script is only slightly faster. It would make sense that the refactored script is faster since the original script would go through every row on the worksheet even if they weren't related to the current ticker for every ticker, and the script would keep switching sheets over and over to add data. The refactored script optimizes the process by creating variable arrays and only looping through the data once. Even with the additional time added to formatting the code, the optimizations are enough to improve the speed of the script by at least .02 seconds in the samples shown above. More runs would need to be made to see whether this improved speed remains consistent, but the formatting helps in improving the visualization of the data. The refactored code also seems more consistent with both years taking the exact same amount of time but this might be coincidental.

## Summary

- What are the advantages or disadvantages of refactoring code?

Refactoring code gives a second chance to look over the code to find optimizations and functionality. Often, the advantages of refactoring can help improve the speed at which the code runs making it take less time, or add features that were not in the scope of the original code. Refactoring can also help improve the readability of the script, or help clarify certain sections of code. Unfortunately, sometimes refactoring can have a negative effect of increasing the complexity of the code, even if the additional complexity can help improve the code. Another disadvantage of refactoring could be understanding the unknowns, trying to gauge how which areas to improve and the difficulty, which could take time and add to feature creep.

- How do these pros and cons apply to refactoring the original VBA script?

When refactoring the original VBA script, there was a definite increase in complexity by adding arrays and a nested loop, as well as including formatting to help improve the visualization of the data output. One of the keys to the refactored script was taking advantage of how the original green stocks data was set up, with all tickers being sorted alphabetically and by date. Since tickers were sorted alphabetically, the script could just move to the next ticker without having to rescan the entire column. This means that scope that the refactored script was working in was more narrow than the original script, since if the tickers weren't alphabetized, then the old script would still work while the new script would not. Other than that, the new script improved the old script by reducing the number of times scanning the entire column would need to happen, as well as having the additional benefit of improving the visualization of the output. In the future, more refactoring could be made to potentially improve the flexibility of the code so that the data could be cured if not in an ideal state as well as adding more functionality to the code.
