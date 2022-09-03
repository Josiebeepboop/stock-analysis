# stock-analysis
---
## Overview of Project

The client, Steve, prepared a workbook with a macro enabling him to quickly analyze stock data that he has collected. Specifically, it calculates the total daily volumes for different tickers he is following and their returns. This allows him to make informed decisions when investing. In the future, he would like to continue adding on to this data, potentially even adding new tickers. However, he is concerned that in the long run, the analysis will take too long. He would like us to enhance the code he has already written so that it runs faster, smoother, and for a larger dataset.

---
## Purpose

To refactor the client's macro code enabling it to run faster and more efficiently for a larger dataset.

---
## Analysis and Challenges

The spreadsheet provided: [Stock Analysis](VBA_Challenge.xlsm)

---
### Results

The original code ran in 0.8046875 seconds for 2017 and 0.78125 seconds for 2018. Looking through the code, it was problematic in a few ways:

1. The code had no clear organization, which made it difficult to read. It started with declaring some variables, then moved on to formatting the output sheet, before . The code was cleaned up to declare all values and variables upfront, followed by the necessary calculations, and ended with formatting of the output sheet. 

2. The code was littered with magic numbers. To avoid this, variables were created to define the numbers and continue to ease readability. In addition, by adding variables instead of hard coded numbers and declaring them upfront, the macro can 

![Refactored 2017 Analysis](/Resources/VBA_Challenge_2017.png)

![Refactored 2018 Analysis](/Resources/VBA_Challenge_2018.png)

---
## Summary

Refactoring code helps to make the code run faster by simplifying it and making it easier to read/understand. This helps with continuity and longevity. It can however be very time-consuming to complete, especially when the code is long and complicated. It can also be risky, especially when unfamiliar with the data and the purpose of the code. It is always best to have clear lines of communication when refactoring.

---
