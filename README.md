# Stock Analysis with VBA and Excel
Performing analyisis on stock market data with a refactored VBA code.

## Overview of the Project

### Purpose
For this challenge we are refactoring an already established VBA macro that loops through stock market data. We will then determine whether this refactored code is more efficient at running through the data. The original code has proven to work for dozens of stocks, however it might not work well for thousands of stocks. The goal of refactoring the code will be to execute analytics on a larger dataset of the stock market with a faster VBA script. 

### Results

Stock market data was analyzed for 2017 and 2018, with our code outputting total volume and return percentage for each ticker (ie each stock). The original code was refactored to loop through the stock data once and collect all of the information efficiently. The code was already set up to run the data based on the year (2017 or 2018) and formatted to output the data on the “All Stocks Analysis” sheet. Header rows were established, and an array of tickers (each stock) were formatted within the code so the focus was on ensuring the for loops were refactored. Below is the refactored code with descriptions included for each step. 




### Summary

Refactoring is an important step in the coding process as it can improve the code. Refactoring is not intended to add functionality or features to the code, but it can lead to a cleaner more organized code. Often times we are looking for the fastest and most efficient way to achieve our goals, so refactoring a code to make it run faster and cleaner has huge advantages. Additionally a cleaner more organized code can be easier to read which helps in understanding the logic as well as simplifying any debugging. Disadvantages of refactoring include potential lengthy time consumption spent that could be used on further development of other code, especially if you are working on a rather large code. Also if a code is already stable then refactoring might not be the best solution.

Our goal for refactoring our stock analysis code was to reduce the amount of time it takes for the macro to run. We were successful in refactoring the code to produce a shorter run time. Both the original code and the refactored code produced the same results and were simple to understand, but having a faster code will allow us to quickly analyze a larger data set. We were able to produce a simpler and more effective code by looping through the data once to get our outputs instead of the nested loops that were in the original code. Below are screen shots of the time run for the original code and the refactored code.

2017 and 2018 Original Code


<img width="262" alt="2017 Run Time" src="https://user-images.githubusercontent.com/89098766/133675816-da88e14d-c283-4de0-b7bf-839595529bdd.png">

<img width="260" alt="2018 Run Time" src="https://user-images.githubusercontent.com/89098766/133676095-a29e0e3b-c7cb-4536-a956-91eb1a97f2b0.png">



2017 and 2018 Refactored Code


<img width="264" alt="2017 Refactored" src="https://user-images.githubusercontent.com/89098766/133676126-8a2ec014-e3f9-40f9-a610-104c06eedb2b.png">

<img width="258" alt="2018 Refactored" src="https://user-images.githubusercontent.com/89098766/133676143-62de8b8f-294b-498b-9010-7f70e522323a.png">



Output for both the original and refactored codes.


<img width="245" alt="All Stocks Analysis 2017" src="https://user-images.githubusercontent.com/89098766/133676396-4e7975a6-1758-49e8-a49a-fa5a5fb5cf4c.png">

<img width="244" alt="All Stocks Analysis 2018" src="https://user-images.githubusercontent.com/89098766/133676442-095ed647-0acc-49d9-b51f-c296de159f4c.png">


