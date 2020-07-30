# VBA of Wall Street
## Overview
### Purpose
The purpose of this analysis is to assist Steve in choosing stocks for his parents to buy based on previous data from 2017 and 2018.

### Background
We had previously assisted Steve by creating a Workbook that would allow him to quickly sort through the data he had created. This worksheet had code that was a bit clunky and took a while to run. Specifically, the first script took 1.09375 seconds to run for the 2017 data set and 1.074219 seconds for the 2018 data set. Our goal is to refactor the code so that it may run faster and have to ability to run more data efficiently. This would assist Steve for future data sets and make it more efficient when he expands his data set to the entire Stock Market.

## Analysis and Challenges
### Analysis
- Steve is attempting to find what stocks his parents should invest in based on the data from 12 companies 2017 and 2018 returns.
- In order to assist Steve, I created a Worksheet that includes a table that finds the return of each stock per year. In this table we can see based on conditional formatting added which stocks have a positive rate of return and which stocks have a negative rate of return based on the color of the cell. Positive returns have a green cell and negative returns are colored red. 
- Based on the data, Steve should recommend that his parents invest in the ENPH and RUN stocks because they are the only two stocks listed that had a positive return in both 2017 and 2018.

![2017 Stock Data](https://github.com/AnnieShaffer/stock-analysis/blob/master/Resources/All_Stocks_2017_Data.png)

![2018 Stock Data](https://github.com/AnnieShaffer/stock-analysis/blob/master/Resources/All_Stocks_2018_Data.png)

### Refactoring
- Steve then let us know that he wanted to look at the whole stock market for these two years. To do this, I refactored the initial script so that it would be able to run more data in less time. 
- In the original code, we nested the loops to find the total daily volume, starting prices, and endings prices, of each stock. In order to refactor the code I set the index to 0 before going further. Additionally, I created a stand alone loop to set the volume of each stock to zero so that Excel did not have to go back and do it for each loop. I also created a loop that did not switch sheets multiple times to take out redundancies.
- Together these factors created efficiencies before starting the loop that would find the total volume, starting prices, and ending prices for each stock because they did not have to run as many times. These efficiences were able to run the data in about a tenth of the time.

![2017 VBA Data](https://github.com/AnnieShaffer/stock-analysis/blob/master/Resources/VBA_Challenge_2017.png)

![2018 VBA Data](https://github.com/AnnieShaffer/stock-analysis/blob/master/Resources/VBA_Challenge_2018.png)

### Challenges
Some of the challenges I encountered during this analysis occurred in writing the code for the macro. When attempting to test the macro to ensure everything ran smoothly, I was receiving and overflow error that I could not figure out. With research I found that I had created a loop for the tickerVolumes that was not needed. This was causing there to be an overflow of information into the return formula. Additionally, I did not initially add the index correctly and had to go back into the code and add it where necessary for the macro to run.

## Summary
### What are the advantages or disadvantages or refactoring code?
One of the advantages of refactoring code is that the code should end up running much faster once it is complete. Additionally, if the program has to run through data fewer times, it is able to take in more data. This means that if code is refactored to be more efficient, more data can be inputted without the program crashing or freezing because it is overwhelmed.

A disadvantage of refactoring code is that it takes a lot of time. If I had faced no challenges refactoring this code, it still would have taken me quite a bit of time to complete. When I ran into challenges it took me several hours over a couple of days to figure out why the code would not run properly. 

### How do these pros and cons apply to refactoring the original VBA script?
The pros applied when removing the loops made the script much quicker as shown in the photos used for the analysis. It also made the script much easier to read if a colleague were to go in and refector the script further.

The cons occurred while actually refactoring the script and are not evident when simply looking at it from an outsiders perspective. However, I felt a lot of frustration when refactoring that I cannot quantify.
