# An Analysis of Stock Trends
---
## Performing analysis on Stock Data to uncover trends
---
### Stock Analysis on trends from 2017 and 2018
---
#### Background of Project
---
##### Analysis of different stocks was performed to help Steve’s parents find the best investment for their money. They were going to put their money in the “DQ” stock and Steve wanted to ensure his parents were doing the right thing. Going beyond the “DQ” analysis, the entire stock market was compared over the last few years. The original code was based only on 12 stocks in the market. The code was then refactored to ensure if thousands of stocks were analyzed then the code would still work well. Refactoring code is a key part in ensuring codes are more efficient and easier for future users to read. 
---
#### Analysis of original code
---
##### To start off the analysis, “DQ” stock was analyzed to see how well it has been performing over the last few years, specifically 2017 and 2018. First, the subroutine was created to hold the code, and then the worksheet was activated so VBA knows where to run the code. A range was created to outline the range of where DAQO stock information can be found in the original worksheet tab, and then the worksheet was formatted with a header row. Then the stock data from 2018 worksheet was activated to be able to loop through the data. Variables were created: “startingPrice” and “endingPrice” and set as double, meaning it’s a data type that allows for decimals. The starting values were set to zero before the initial loop through the data. This part of the code is below:
![image](https://user-images.githubusercontent.com/98365963/158487654-e77497c6-00b6-41a9-adf2-6f657c8ca79e.png)
##### Now it was time to create a loop to run through all the data at once and summarize the data. First step was to establish the number of rows to loop over, and this was done as follows:
![image](https://user-images.githubusercontent.com/98365963/158487708-f987b73b-735e-4e2e-afef-4bf8125cee90.png)
##### Now it was time to loop through the data to summarize everything. The start of the for loop tells excel where the start of the row is and where it ends. This is useful and helps if more information is added later in the process. This code was done as follows: ` For i = rowStart To rowEnd.` Now if conditions were set to summarize the data. The first if condition was to increase the totalVolume, which was set to zero before the for loop, and look at the current value in the cell. If the current value in the cell belonged to the current ticker, or stock label, then that value became the new totalVolume. This is shown in the following code: 
![image](https://user-images.githubusercontent.com/98365963/158487779-2ca7f63a-caf6-4b0a-b662-31613261e4f5.png)
##### The next if condition was coded to look for where the price of the stock started in the data. The following if condition checks the cell it’s in, and then looks at the previous cell to see if it belongs to a different ticker. If the above cell belongs to another ticker, then the current cell is labeled as the startingPrice variable. This is shown in the following code: 
![image](https://user-images.githubusercontent.com/98365963/158487812-6052c6fa-46dc-4f2c-9ed1-9d8058512d28.png)
##### The next if condition was coded to look at the cell below the current cell to look for a different ticker value. If the cell below had a different ticker, then the current cell is the endingPrice for that stock ticker. This is shown in the following code: 
![image](https://user-images.githubusercontent.com/98365963/158487850-5f28686e-5065-4c6e-afd6-3fb36f186147.png)
##### Now that all the if conditions were coded out, then the for loop ends with the statement: `Next i` so the loop can start again at a different row to continue looking at all the data. Lastly the data is summarized in a new worksheet tab called “DQ Analysis” and the data is summarized under the headers created earlier in the code and the subroutine is completed. This is shown by this: 
![image](https://user-images.githubusercontent.com/98365963/158487886-1202ade4-8cf2-4fde-a035-4f7fc15c0e24.png)
##### The above code was refactored to loop over all the data for the remainder of the stocks in the data. The only changed that were made were more variables and tickers were added, as shown here: 
![image](https://user-images.githubusercontent.com/98365963/158488001-f4a04ecb-56f5-442e-b1b8-88584f4a367e.png)
##### Then everywhere “DQ” was in the code, “ticker” was placed instead. This was to ensure that all of the stocks were looped over and then summarized into a new worksheet labeled “All Stocks Analysis.” This is shown here: 
![image](https://user-images.githubusercontent.com/98365963/158488045-b496f7d1-79bd-481c-b7ba-98d5e3cf823b.png)
##### After going through this analysis, the above code for the “All Stocks Analysis” worksheet was refactored to be able to handle more information as more stocks were added to the data. 
---
#### Analysis of Refactored code
---
##### Per the analysis using the refactored code, the only changes from the original code are arrays were added to aid in the analysis process later down the line. As seen here: 
---
#### Summary
---
##### Advantages of refactoring code are it can make the code run quicker and easier, with less steps; also, it makes it easier for the next person to read and understand it; and finally making it efficient while using less memory on your computer. Another advantage to refactoring code is making it better so when more data is added, then the code doesn’t have to be changed. The pros of refactoring code apply to the original VBA script by simplifying the steps, with added comments it did make it easier to read, and it made it efficient to where the code ran faster than the original script. The disadvantages of refactoring code are you really must understand the original code to be able to rewrite and shorten it. A challenge I faced with this module was understanding the code enough to change it and make the variables into arrays. I was struggling with understanding where to use “tickerIndex” and how to even use it to make the code efficient. 
