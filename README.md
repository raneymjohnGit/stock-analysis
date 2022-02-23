# Stock Analysis with VBA
    ** This is a repository to store the Module 2 VBA Challenge **

## Overview of Project
Steve wants to do a little more research for his parents, he wants to expand the existing green_stocks dataset to include the entire stock market over the last few years.  Although the code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute. Steve wants to make the solution code much simpler and run faster.

### Purpose
In this challenge, we need to refactor the Module 2 solution code to loop through all the data one time in order to collect the same analysis that we did in this module. Then determine whether
refactoring the code successfully made the VBA script run faster or not. Refactoring means: you are
trying to make the code more efficient by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. 

## Analysis and Challenges
Deliverable 1: Refactor VBA Code and Measure Performance
    
    1. Download the challenge_starter_code.vbs file and rename it VBA_Challenge.vbs.
    
    2. Create a folder called "Resources" to hold the run-time pop-up messages that you will screenshot after running refactored analyses for 2017 and 2018.
    
    3. Rename the green_stocks.xlsm file that you used in this module as VBA_Challenge.xlsm.
    
    4. Add the VBA_Challenge.vbs script to the Microsoft Visual Basic editor.
    
    5. Use the steps below to add code where indicated by the numbered comments in the starter code file.
        
        Step 1a: Create a tickerIndex variable and set it equal to zero before iterating over all the rows.
        
        Step 1b: Create three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices. The tickerVolumes array should be a Long data type. The tickerStartingPrices and tickerEndingPrices arrays should be a Single data type.

        Step 2a: Create a for loop to initialize the tickerVolumes to zero.

        Step 2b: Create a for loop that will loop over all the rows in the spreadsheet.
        
        Step 3a: Inside the for loop in Step 2b, write a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker. Use the tickerIndex variable as the index.
        
        Step 3b: Write an if-then statement to check if the current row is the first row with the selected tickerIndex. If it is, then assign the current starting price to the tickerStartingPrices variable.
        
        Step 3c: Write an if-then statement to check if the current row is the last row with the selected tickerIndex. If it is, then assign the current closing price to the tickerEndingPrices variable.
        
        Step 3d: Write a script that increases the tickerIndex if the next row's ticker doesn't match the previous row's ticker.
        
        Step 4: Use a for loop to loop through your arrays (tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices) to output the "Ticker", "Total Daily Volume" and "Return" columns in your spreadsheet.
    
    5. Finally, run the stock analysis, then confirm that your stock analysis outputs for 2017 and 2018 are the same as they were in the module (as shown in the images below). In your Resources folder, save the pop-up messages showing elapsed run time for the refactored code as VBA_Challenge_2017.png and VBA_Challenge_2018.png.


### Analysis of All stocks for the year 2017
![image_name](https://github.com/raneymjohnGit/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

#### Time taken for Refactored code 
![image_name](https://github.com/raneymjohnGit/stock-analysis/blob/main/Resources/Refactored_VBA_Challenge_2017_TimeTaken.png)

#### Time taken for Original code 
![image_name](https://github.com/raneymjohnGit/stock-analysis/blob/main/Resources/Green_Stocks_2017_TimeTaken.png)

### Analysis of All stocks for the year 2018
![image_name](https://github.com/raneymjohnGit/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

#### Time taken for Refactored code 
![image_name](https://github.com/raneymjohnGit/stock-analysis/blob/main/Resources/Refactored_VBA_Challenge_2018_TimeTaken.png)

#### Time taken for Original  code 
![image_name](https://github.com/raneymjohnGit/stock-analysis/blob/main/Resources/Green_Stocks_2018_TimeTaken.png)

### Challenges and Difficulties Encountered

None

Reference documents:
1.  Module 2 from our Boot Camp


## Results
By looking at the images, Results from the refactored code (VBA_Challenge) and module 2 solution code (Green_Stocks), the results are identical. However, the execution time for the refactored code is much less compared to the module 2 solution code.

## Summary
1. What are the advantages or disadvantages of refactoring code?
    
    Advantages:
    
    i. When refactoring code, you are not adding new functionality; you just want to make the code  
        more efficient by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. 
    
    ii. Refactoring is common on the job because first attempts at code won't always be the best 
    way to accomplish a task. 
   
   Disadvantages:

    i. When refactoring code, you may need to spend extra time to make it work exactly like the original code. 
    
    ii. Prone to bugs and errors, if not did it due diligently. 

2.  How do these pros and cons apply to refactoring the original VBA script?

    i.  Took some time to figure out how this refactoring should work.

    ii. Initially got few errors before i make it work.


