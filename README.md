# Stock-Analysis
Module 2
## Project Overview
The original project was to use VBA to create an Excel spreadsheet to analyze a dozen stocks, providing data on total daily volume and yearly return for each stock, to help a client determine whether particular stock(s) were worth investing in. The purpose of this current project was to refactor the original code, making it both faster and more efficient, so the dataset could be expanded to include a larger number of stocks over an extended number of years.
## Results
In the original code, the execution of the nested loop used a lot of run time because the program ran repeatedly for each stock:
When refactoring the code, I changed the code, using if-then statements, to run through the entire worksheet rather than looping one stock at a time:
[GitHub Pages] https://github.com/HollyESmith/stock-analysis/commit/ea7e4daeafa42e456e77dd47143d4a4fe7a32e41#diff-05347f5b98788640ecd4c7a17aee263d91186f5f5350d21ee71954f4e6fa4d49
### Refactoring Code
Consequently, the refactored code run time for 2017 data was almost 8 times faster [7.916] than the original code: 

The refactored code run time for 2018 data was almost 7 times faster [6.724] than the original code:

### Stock Performance
The stocks analyzed in this project were highly volatile. There were only 2 of 12 stocks that had a positive net return in both 2017 and 2018: ENPH & RUN. However, over this two-year period, ENPH and SEDG were the best performers. This chart shows stock prices based on a hypothetical beginning value (at 2016 end) of $100:
## Summary
Advantages of refactoring code include efficiency, readability, and reproducibility – that is, when the next coder understands the cleaner code, he or she will be able to continue to refine/refactor the code as well as re-use pieces of the code for other projects. Disadvantages of refactoring code may include 1) it’s a time-consuming process and 2) “if it ain’t broke, don’t fix it” – the coder could end up spending an inordinate amount of time debugging the intended improvements.
