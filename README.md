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
The VBA code written for the **_YearAllStockAnalysis_** module and the refactored code for the **_AllStockAnalysisRefactored_** module produces the same spreadsheet output as shown in **_Fugure 1_** along with the timer pop ups shown in **Figures 2 & Figures 3**.  

**_Figure 1: Spreadsheet results refactored module_**

![Spreadsheet results](/Year_Over_Year_Comparison.png)

**_Figure 2: Performance Results 2017_**

![2017 Timer Compare](/2017_Comparison_Orig_vs_Refact.png)

**_Figure 3:Performance Results 2018_**

![2018 Timer compare](/2018_Comparison_Orig_vs_Refact.png)

### Changes to the Index Array

The original array has its value's manually set.  If the source data were to change the code would need to be updated to reflect any added or removed tickers.

**_Figure 4: Original Manual Setting of Array_**

![Original Array code](/initializing_array_for_all_tickers.png)

In the refactored code the ticker array is created dynamically. 

**_Figure 5: Refactored Dynamic Array_**

![revised Array code](/ticker_index_from_dictionary.png)

**_Figure 6: Dictional to Create Unique List of Values for the the Array_**

![special function dictionary](/FunctionGetUniqeNames.png)

### Changes to the Looping Pattern & Use of Arrays

**_Figure 7: Original For Loop & Embedded loop_**

![original embedded for loop](/original_code_nested_for_loop.png)

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





