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
In this section w

*Figure 1: Performance Results 2017*

![2017 Timer Compare](/2017_Comparison_Orig_vs_Refact.png)

*Figure 2:Performance Results 2018*

![2018 Timer compare](/2018_Comparison_Orig_vs_Refact.png)

### Changes to the Index Array

The original array has its value's manually set.  If the source data were to change the code would need to be updated to reflect any added or removed tickers.

*Figure 3: Original Manual Setting of Array*

![Original Array code](/initializing_array_for_all_tickers.png)

In the refactored code the ticker array is created dynamically. 

*Figure 4: Refactored Dynamic Array*

![revised Array code](/ticker_index_from_dictionary.png)

*Figure 5: Dictional to Create Unique List of Values for the the Array*

![special function dictionary](/FunctionGetUniqeNames.png)

### Changes to the Looping Pattern & Use of Arrays

*Figure 6: Original For Loop & Embedded loop*

![original embedded for loop](/original_code_nested_for_loop.png)

*Figure 7:Refactored Single Loop*

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





