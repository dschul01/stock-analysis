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
The initial code was enhanced using an index for all the ticker's performance measurement loops and resulted in significant time reductions to run the code as seen in the images below.  The time differential would be even greater if the volume of stocks to analyze increased.
![Refactored_Time_Impacts.png](https://github.com/dschul01/stock-analysis/blob/main/Refactored_Time_Impacts.png)
## Summary
### Advantages of Refactoring Code
There are several advantages of refactoring code.  It leads to better quality code making it more manageable and cleaner for the original programmer or others who might need to update it in the future.  The process of refactoring will also help in finding bugs which might not have originally occurred when running initial use cases.  It also often reduces run times as it leads to removing unnecessary code.  

The refactoring for this particular project resulted in all the advantages listed.  The original code had unnecessary lines of code which made the performance run over 6 times longer before refactoring.  The new code is cleaner and runs much quicker as seen in the Results above.
  
### Disadvantages of Refactoring Code
There are also disadvantages of refactoring code.  Refactoring takes time and money. A programmer must go back through the code deciphering what it's doing and then use resources to refactor.  There are possibilities refactoring does not improve timeliness and could break other applications which depend on the source code being adjusted.

The refactoring for this particular project resulted in the use of additional hours spent to create better quality code.  However, the resulting 6x time savings made it a worthwile endeavor.


