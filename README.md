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
  
## Refractoring Results
The initial subroutine looped through each of the stock tickers and rows. By creating output arrays and initializing the ticker volume to zero, we could create a more efficienct looping system. These creates more space in memory as it loops. 
<p align="center">
<img src="https://github.com/teachjanderson/stock-analysis/blob/main/images/Initializing.png" width="600" />

This next step was a challenge, and took some research to solve (which is referenced at the end of the analysis). As the subroutine looped through the tickers, it needed to identify when one ticker ended and another began. Using the following code, the subroutine identified as it looped through a new ticker name. The loop identifided when the ticker name changed and went back one to close and begin a new correct loop. While this idea makes sense, refractoring the code proved challenging. In the end, this proved an efficient method. 

  For i = 2 To RowCount

        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                tickerStartingPrice(tickerIndex) = Cells(i, 6).Value
            End If
        
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerendingPrice(tickerIndex) = Cells(i, 6).Value
        
                tickerIndex = tickerIndex + 1
            End If
        
        Next i

<p align="center">
<img src="https://github.com/teachjanderson/stock-analysis/blob/main/images/Looping.png" width="600" />

Finally, another loop to output the values onto the spreadsheet and adding buttons to make it simple for the user. It also resulted in a faster run time. This refractored code had a run time of about 1/10th of a second, as seen in the images below. 
<p align="center">
<img src="https://github.com/teachjanderson/stock-analysis/blob/main/images/Looping2.png" width="600" />
  
<p align="center">
<img src="https://github.com/teachjanderson/stock-analysis/blob/main/images/2017%20Refrectored.png" width="600" />
<p align="center">
<img src="https://github.com/teachjanderson/stock-analysis/blob/main/images/2018%20Refractored.png" width="600" />

This refractoring resulted in a change of nearly 85% in comparison to the original time. As the data analyzed is increased, this will prove more beneficial. 
<p align="center">
<img src="https://github.com/teachjanderson/stock-analysis/blob/main/images/refractortime.png" width="600" />

  
  
## Summary
**What are the advantages or disadvantages of refactoring code?**
There are several advantages to refracting code. The first is efficiency or performance. In this case, while the human use is not noticeable at this scale, the refractoring increased performance dramatically as represented by percent of change. The second reason is understanding of the code. The refractored code did not increase or decrease the clearness of the code, but in general refractoring towards simplicity and understanding (readability) is always an improvement. Two disadvantages of refractoring code are time and complexity. Understanding the original script, identifying ways to improve while maintaining clarity, takes time and effort. It's quite possible all that time makes few if any improvements. 

**How do these pros and cons apply to refactoring the original VBA script?**
In our case, referencing the points above, there was a substantial change in proficiency and that made the refractoring a pro. However, this did, as mentioned, come with time requirements and understanding. There were complexities and challenges in understanding limitations of the original script and identifying ways to improve them. 
  
**References**
In initially refractoring the code, I altered my methods considerably in running into challenges with the If and Else statements. After analyzing and understanding different approaches and solutions to other problems, I found this method to identify when the tickers changed. This led me to a greater understanding of the loops as well and more efficient methods using our output arrays. One pro of this entire refractoring is a better understanding of the original code, memorizing basic syntax, and understanding efficient ways to loop through large amounts of data. 
https://zoom.us/rec/play/3y8QSEzGpawN1aFaIdTTB07ORUoGVwlzCxrhpsY_vSxIxMMKwIEolwHZrynFP_TXJ1XLWYVjWcEC4rxd.4q3FJMwSSyZxC3BG?startTime=1637415343000&_x_zm_rtaid=-i-cEITxScOE-OFMLjGzqQ.1637430207851.2075cf95887cdc9463ec29789c0911c9&_x_zm_rhtaid=958
