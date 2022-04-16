# Stock Analysis Using Excel and VBA
## Overview of Project
The purpose of this project was providing our client with accurate and efficient information on the returns of 12 stocks for a given year. To accomplish this, a VBA Macro was created with buttons to run the data. To increase the effectiveness, the code was refractored to improve output time. While the current sheet runs roughly 3000 rows of data, improving the code creates the possibility of running larger data sets at an efficient speed. The main goal of this process was increasing speed through altering the code to more effectively use memory. 

The Macro returns three pieces of information: the different stocks by name, their total volume, and the yearly return as a percent. The name of the stocks is reported by abbreviation, the daily volume is the shares traded and the return is the percent different from the starting price to ending price in a given year. 

## Results
Initially, a subroutine was created to analyze the performance of stocks for two given years: 2017 and 2018. While many stocks showed positive performance in 2017, only two stocks, ENPH and RUN continued showing positive performance in 2018. 
<p align="center">
<img src="https://github.com/teachjanderson/stock-analysis/blob/main/images/StockAnalysis.png" width="600" />

As seen below, these initial subroutines took over 6 tenths of a second to complete. While this is a blink of an eye, increasing the speed by refractoring the code provides the efficiency as the number of stocks are analyzed. This dataset was reasonably small compared to one 10 or 100 times its size. Therefore, the initial subroutine was refractored to increase efficiency in memory and speed. 
<p align="center">
<img src="https://github.com/teachjanderson/stock-analysis/blob/main/images/2017.1.png" width="600" />
  
<p align="center">
<img src="https://github.com/teachjanderson/stock-analysis/blob/main/images/2018.png" width="600" />

  
Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
Summary: In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?
