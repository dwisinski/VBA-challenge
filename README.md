# VBA-challenge
by: Dave Wisinski

February 2021

## Project: VBA Homework - VBA of Wall Street

## Description:
This project was created to demonstrate a number VBA scripting concepts on an actual data set of historical stock market data. Concepts covered include: conditionals, loops, nested loops, conditional formatting, automated data analysis and calculation, etc.

## Notes:
The "stockData" subroutine contains the main script for analyzing the initial data set and reading, storing and performing calculations on various values, then inserting stored values/calculation results into a separate table. This subroutine also contains some conditional formatting instructions at the end.

The separate "getAdditional" subroutine runs automatically (via the "Call" instruction at the end of the stockData sub) after the initial sub, creates an additional table of data calculations, and also performs some minor formatting.

## **Important Usage Notes:**
The VBA script file (VBA-challenge-script.bas) should be run on the provided data set (Multiple_year_stock_data.xlsm). 

**Only the "stockData" sub should be run as it automatically calls the getAdditional sub as noted above.**