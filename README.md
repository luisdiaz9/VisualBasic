## VisualBasicHardAllSheets
### Summary
This repository uses VBA scripting to analyze real stock market data.<br>
The analysis shows the Yearly Change, Percent Change and Total Stock Volume per Ticker.<br>
The information is carved up annually in different sheets and the code runs all of them at once.<br>
### Technical Details
In order to run the code, it is required Microsoft Visual Basic.
#### Files:
Test Data - Use this while developing scripts.<br>
Stock Data - Run scripts on this data to generate the final homework report.<br>
Stock market analyst.<br>
### Screenshots
2016.JPG<br>
![Output](2016.JPG)
### Explanations
The outcome is shown in screenshots for reference purpose of the public.<br>


# VisualBasicHardAllSheets
The VBA of Wall Street<br>

Easy<br>
Create a script that will loop through one year of stock data for each run and return the total volume each stock had over that year.<br>
Display the ticker symbol to coincide with the total stock volume.<br>

![Easy](easy_solution.png)

Moderate<br>
Create a script that will loop through all the stocks for one year for each run and take the following information.<br>
The ticker symbol.<br>
Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.<br>
The percent change from opening price at the beginning of a given year to the closing price at the end of that year.<br>
The total stock volume of the stock.<br>
You should also have conditional formatting that will highlight positive change in green and negative change in red.<br>
The result looks as follows.<br>

![Moderate](moderate_solution.png)


Hard<br>
The solution includes everything from the moderate challenge and is able to return the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume".<br>
Solution looks as follows.<br>

![Hard](hard_solution.png)


The script allows it to run on every worksheet, i.e., every year, just by running it once.<br>

Use the sheet alphabetical_testing.xlsx while developing the code. <br>
This data set is smaller and allows to test faster.<br>
The code should run on this file in less than 3-5 minutes.<br>
