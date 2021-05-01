# VBA_Challenge
 
## Overview of Project

### Purpose

Here we are set to the task of refactoring VBA code to collect stock data in the years 2017 and 2018. Once the data is run through our code the user will be able to swiftly see which stocks are worth their investment. Our goal was to refactor the original code to make the process run quicker.

### Our Dataset

From the 12 data sets we were presented with (ticker value, stock issue date, opening, closing and adjusted closing price, highest and lowest price, volume of stock) the goal was to analyse the ticker, total daily volume, and the returm on each individual stock.

## Results

### Analysis

To begin teh refactoring process I first needed to copy the code from the previous macro to create the chart headers, input box, ticker array and worksheet activation. Guidance for refactoring were set in order as instructions. The main difference between the two sets of code is that the refactored code does not include a nested For loop. Instead within a Loop, three separate If statements take the place of the nested loop on the unfactored code. There are more lines of code, however the process runs almost a full second quicker.

## Summary

### Benefits and Downfalls of Refactoring Code for Stock Analysis

Although the total lines of code is greater, the code itself is cleaner - benefiting from the lack of nested For loops - and runs quicker. A code that is easier to read and follow, allows for any future programmer to understand with greater ease what the code is supposed to be doing. Overall the greatest benefit as a result of refactoring was the decrease in macro run time. Originally the analysis took ~1 second to run, our refactored analysis takes about 1/4 of that time. The refactored runtimes are detailed below:



###2017

![VBA_Challenge_2017.PNG](https://github.com/craig-clemens/VBA_Challenge/blob/main/Resources/VBA_Challenge_2017.PNG)

###2018

![VBA_Challenge_2018.PNG](https://github.com/craig-clemens/VBA_Challenge/blob/main/Resources/VBA_Challenge_2018.PNG)