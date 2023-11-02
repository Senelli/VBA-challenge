# VBA-challenge

* This VBA script goes through all the daily data in each worksheet and calculates and displays the 
    - yearly change calculated from finding the difference between opening value at the beginning of the year and the closing value at the end of the year,
    - percent change from the opening price at the beginning of the year to the closing price at the end of the year, 
    - and the total stock volume at the end of the year
  for each stock/ticker. Then displays the data along with their respective stock/ticker on a new separate data table.
* Conditional formatting applied to show positive change in green and negative change in red.

* Then it goes through the new data table to find the stock/ticker with the greatest increase in percent, the greatest decrease in percent, and the greatest total stock volume and their respective stock/ticker and displays this data in a new data table. This data is calculated for each year (each worksheet).