# Spread-sheet-Processor
Python program which takes in values in a set column, applies a 10% discount and creates a chart to show them off.

The program consists of a simple function “process_workbook” that takes in the name of the file then it proceeds to:
1.	Load the workbook.
2.	Access the Sheet that the data is stored inside.
3.	Loop through rows 2 to 4 inclusive.
4.	Access the individual cells at the third column.
5.	Multiply their values by 0.9 to apply a 10% discount.
6.	Store the discounted values in the fourth column.
7.	Create a reference of the values in the fourth column.
8.	Create an instance of the BarChart class.
9.	Feed the chart the data from the fourth column.
10.	Create a bar chart using the data from the fourth column with it’s upper left corner at cell “e2”.

