# Stock Analysis Using VBA
Performing analysis on energy stock data over multiple years while refactoring initial code to increase efficiencies.
## Overview of Project
### Purpose
The purpose of this project was to utilize VBA to automate the task of analyzing a number of energy stock's performance over 2 years to help determine if they are worth the investment.  Further steps were taken to identify ways to refactor, or improve the code by making small changes without changing the output.  The file, [VBA_Challenge.xlsm](https://github.com/dschul01/stock-analysis/blob/main/VBA_Challenge.xlsm), contains the original and refactored code as well as UI buttons to calculate each year's analysis.  The output will also provide the time to run through the macro for the user to see how well the refactored code improves efficiencies.

## Results
### Stock Performance Using Loops
Twelve stocks were analyzed between the years 2017 and 2018.  The key objective was to automate the analysis rather than manually entering formulas.  Examples are provided which show the effectiveness of creating code to calculate stock performance measurements utilizing loops rather than using Excel functions which would have to be added as more stocks are evaluated.
![Loop_Measurement.png](https://github.com/dschul01/stock-analysis/blob/main/Resources/Loop_Measurement.png)

There are two stocks, ENPH and RUN, over the period which are revealed as providing positive results YoY.  The code seen below makes it quite easy to visualize by highlighting the positive returns in green seen in the output.
![Formatting_Code.png](https://github.com/dschul01/stock-analysis/blob/main/Resources/Formatting_Code.png)
![Positive_Returns_YoY.png](https://github.com/dschul01/stock-analysis/blob/main/Resources/Positive_Returns_YoY.png)
### Refactoring Impacts
The initial code was enhanceD using an index for all the ticker's performance measurement loops and resulted in significant time reductions to run the code as seen in the images below.  The time would be reduced even more as the number of stocks to analyze increases.
![Refactored_Time_Impacts.png](https://github.com/dschul01/stock-analysis/blob/main/Refactored_Time_Impacts.png)
## Summary
### Advantages

### Disadvantages	


