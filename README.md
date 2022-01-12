# Stock Analysis

## Overview of the Project

The code written in module two defines an array of 11 stock tickers and then loops through the data to sum the total volumes and calculate the return. This project focuses on improving the performance of the code written during module two so that it performs more efficiently and uses less resources.  To acheive this outcome it was important to better understand the use of array's, to think creatively about options for changing the code, testing theories, and discarding things that don't work.

### Purpose

Refactoring code a reality in development for a number of reasons some good and others not so much.  Many times code is refactored because the requirements have changed and updates are required.  Other times end users may be experiencing time out errors, capacity constraints, or lengthy processing times frustrating their efforts to complete a task.  Whatever the reason, reviewing and modifying existing code either your own or others is common. This challenge takes the code written in module 2 and asks that the code be improved so that the average run time is decreased and the process loops through the data once.

## Analysis

### Performance Results

- The analysis is well described with screenshots and code


![2017 Timer Compare](/2017_Comparison_Orig_vs_Refact.png)


![2018 Timer compare](/2018_Comparison_Orig_vs_Refact.png)

### Refactored Index Array

The original array has its value's manually set.  If the source data were to change the code would need to be updated to reflect any added or removed tickers.

![Original Array code](/initializing_array_for_all_tickers.png)

In the refactored code the ticker array is created dynamically. 

![revised Array code](/ticker_index_from_dictionary.png)

![special function dictionary](/FunctionGetUniqeNames.png)

### Refactored Loop through the Data & Use of Arrays

![original embedded for loop](/original_code_nested_for_loop.png)

![refactored single loop](/refactored_code_single_for_loop.png)


## Summary

- There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
- There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).
