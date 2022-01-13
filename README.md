# Stock Analysis

## Overview of the Project

The code written in module two defines an array of 11 stock tickers and then loops through the data to sum the total volumes and calculate the return. This project focuses on improving the performance of the code written during module two so that it performs more efficiently and uses less resources.  To acheive this outcome it was important to better understand the use of array's, to think creatively about options for changing the code, testing theories, and discarding things that don't work.

### Purpose

Refactoring code a reality in development for a number of reasons some good and others not so much.  Many times code is refactored because the requirements have changed and updates are required.  Other times end users may be experiencing time out errors, capacity constraints, or lengthy processing times frustrating their efforts to complete a task.  Whatever the reason, reviewing and modifying existing code either your own or others is common. This challenge takes the code written in module 2 and asks that the code be improved so that the average run time is decreased and the process loops through the data once.

## Analysis

In this section we'll review the various parts of the original script which were refactored and why.  The analysis section will cover the following:

- Performance Results
- Changes to the Index Array
- Changes to the Looping Pattern & Use of Arrays

### Performance Results
The VBA code written for the **_YearAllStockAnalysis_** module and the refactored code for the **_AllStockAnalysisRefactored_** module produces the same spreadsheet output as shown in **_Figure 1_** along with the timer pop ups shown in **Figures 2 & Figures 3**.  

**_Figure 1: Spreadsheet results refactored module_**

![Spreadsheet results](/Year_Over_Year_Comparison.png)

Both figures 2 & 3, in the left pop ups show the timer from the original code with a time of 0.59 seconds for 2017 and 0.50 seconds for 2018.  The pop ups on the right show a great improvement in performance with 0.11 seconds for 2017 and 0.09 seconds for 2018.

**_Figure 2: Performance Results 2017_**

![2017 Timer Compare](/2017_Comparison_Orig_vs_Refact.png)

**_Figure 3:Performance Results 2018_**

![2018 Timer compare](/2018_Comparison_Orig_vs_Refact.png)

### Changes to the Index Array

The original ticker array has its values manually set.  If the source data were to change the code would need to be updated to reflect any added or removed tickers.

**_Figure 4: Original Manual Setting of Array_**

![Original Array code](/initializing_array_for_all_tickers.png)

In the refactored code, the ticker array is created dynamically by leveraging a function to read the ticker column into a dictionary and then return an array. This was acheived by refactoring a function which was sourced from https://www.py4u.net/discuss/1443953 answer # 2.  This function, called **_GetUniqeNames_**, uses the dictionary object to return an array from a specified range.  This approach was selected because it would effectively create a unique list of values because the dictionary object in VBA will not allow a duplicate key.  **_Figure 5_** shows the function **_GetUniqeNames_** and **_Figure 6_** Shows the TickerIndex Array being populated by the Function.

**_Figure 5: Dictional to Create Unique List of Values for the the Array_**

![special function dictionary](/FunctionGetUniqeNames.png)

**_Figure 6: Refactored Dynamic Array_**

![revised Array code](/ticker_index_from_dictionary.png)


### Changes to the Looping Pattern & Use of Arrays

In the **_YearAllStockAnalysis_** module the original code uses two *For loops* with the second *For loop* nested to loop through the rows of data.  The outer loop will loop to the first ticker(0) and then will go into the inner loop.  The inner loop does three things at each of the 3013 rows.  First it will total the volume for each row that equals the ticker.  It also determines the starting price and ending price.  Followed by writing the values to the worksheet for the ticker as well as  increment to the next ticker by adding 1.  At this final step the outer loop begins again with the Ticker+1, repeats the cycle until it finishes at Ticker(11).  Each time the outer loop finishes it must write to the worksheet before it can move to the next ticker.  This stop, print, clear the variable and start anew for the next iteration has the effect of making the code run slower.

**_Figure 7: Original For Loop & Embedded loop_**

![original embedded for loop](/original_code_nested_for_loop.png)

The **_AllStockAnalysisRefactored_** module **_Figure 8_** shows the code refactored leveraging the following four arrays:
- tickerIndex
- tickerVolumes
- tickerStartingPrices
- tickerEnding Prices

The For Loop performs the following steps:
1. The counter is set to iterate through the rows
2. The currticker variable is set to the current ticker by the tickerIndex array which is incremented by the tickercounter
3. at the first row, it checks to see if the cells


**_Figure 8:Refactored Single Loop_**

![refactored single loop](/refactored_code_single_for_loop.png)


## Summary

### General Advantages & Disadvantages of Refactoring Code
- There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).

- Advantages
  - Refactoring code ddldldldldld
  - dldldldld
  - dldldldld

- Disadvantages
  - dkdkdkdk
  - dkdkdkd
  - dkdkd


### Advantages & Disadvantages of the Refactored VBA Script
- There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).

- Advantages
  - Refactoring code ddldldldldld
  - dldldldld
  - dldldldld

- Disadvantages
  - dkdkdkdk
  - dkdkdkd
  - dkdkd





