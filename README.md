# Stock-Analysis
Module 2
## Project Overview
The original Module 2 project was to use VBA to create an Excel spreadsheet to analyze a dozen stocks, providing data on total daily volume and yearly return for each stock to help a client determine whether particular stock(s) were worth investing in. The purpose of this current Challenge was to refactor the original code, making it both faster and more efficient, so the dataset could be expanded to include a larger number of stocks over an extended number of years.
## Results
In the original code, the execution of the nested loop used a lot of run time because the program ran repeatedly for each stock:

![Module2_OriginalCode](https://user-images.githubusercontent.com/97558998/155898437-27869ca9-9c61-44d7-a334-e43b156d98e8.png)

When refactoring the code, I changed the code, using if-then statements, to run through the entire worksheet rather than looping one stock at a time:

![VBA_Challenge_RefactoredCode](https://user-images.githubusercontent.com/97558998/155898472-ad41a07a-b0a6-463b-91ca-e4286b17cb48.png)

Consequently, the refactored code run time for 2017 data was almost 8 times faster [7.916] than the original code: 

![VBA_Challenge_2017](https://user-images.githubusercontent.com/97558998/155898594-88b33c45-dbbd-45f5-a45c-ae75c1f5da39.png)


The refactored code run time for 2018 data was almost 7 times faster [6.724] than the original code:

![VBA_Challenge_2018](https://user-images.githubusercontent.com/97558998/155898601-6f926751-9e7c-4bc5-bf4c-bd515b4b651c.png)

#### Stock Performance
The stocks analyzed in this project were highly volatile. There were only 2 of 12 stocks that had a positive net return in both 2017 and 2018: ENPH & RUN. However, over this two-year period, ENPH and SEDG were the best performers. This chart shows stock prices based on a hypothetical beginning value (at 2016 end) of $100:

![VBA_Challenge_StockPerformance](https://user-images.githubusercontent.com/97558998/155898573-cca87e37-8192-4aba-9e25-7c8cb8433c46.png)

## Summary
Advantages of refactoring code include efficiency, readability, and reproducibility – that is, when the next coder understands the cleaner code, he or she will be able to continue to refine/refactor the code as well as re-use pieces of the code for other projects. Disadvantages of refactoring code may include 1) it’s a time-consuming process and 2) “if it ain’t broke, don’t fix it” – the coder could end up spending an inordinate amount of time debugging the intended improvements. These pros and cons applied to this refactoring project. Pros: the refactored code is more efficient as evidenced by the improved run times, and I believe another coder will understand the refactored script and be able to refine it further. Cons: As a beginner coder it was indeed a challenge to refine the script, and I did spend an inordinate amount of time debugging. It was a learning experience.
