# VBA-challenge
VBA scripting to analyze real stock market data.

## Stock Market Analysis Using VBA

A report that analyzes the yearly growth of each stock tickers from 2014 to 2016.

### Background:
The initial stock report contains the stock market volumes, and the opening, closing, highest, and lowest prices per tickers for each business days. Using VBA, we summarized the yearly price changes, percentage of price changes, total stock volumes for each stock tickers, and the tickers which has the greatest and lowest growth by year.

### Analysis:
We created a VBA script that will loop through all the stocks for one year for each run and extract the following information in the summary table.
•	Each unique ticker symbols.
•	The yearly change by subtracting the closing price at the end of the given year with the opening price at the beginning of that year. If there was no opening price, the next non-zero opening price would be assumed as the opening price of the given year.
•	The percent change by dividing yearly change by the opening price. The positive percent change was highlighted in green and the negative percent change was highlighted in red.
•	The total stock volume of each tickers by the given year.
•	The tickers with the greatest percentage growth, lowest percentage growth, and highest total volume in each given year.

 

 

 


