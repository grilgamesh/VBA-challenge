# VBA-challenge
Solution to week 2 homework


This code adds a new sheet at the start of the excel workbook to contain the summary data. It then prints column headings for that sheet.

It then loops through all the sheets in workbook: 
* calculate how many rows long the workbook is
* instantiate the initial opening value for the top row
* loop through the rows, updating total volume, until a change in name is detected, in which case: stop, output the total volume, and calculate the change in price as an absolute value and as a percentage value.

After finishing these loops, the program formats the data correctly.

Then the program calculates the value for the extension task; it loops through the sumamry data once, updating three values as it goes: the largest % increase, % decrease, and total volume. these are then outputted and formatted.