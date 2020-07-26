# VBA of Wall Street
## Overview
### Purpose
The purpose of this analysis is to assist Steve in choosing stocks for his parents to buy based on previous data from 2017 and 2018.

### Background
We had previously assisted Steve by creating a Workbook that would allow him to quickly sort through the data he had created. This worksheet had code that was a bit clunky and took a while to run. Specifically, the first code took 1.09375 seconds to run for the 2017 data set and 1.074219 seconds for the 2018 data set. Our goal is to refactor the code so that it may run faster and have to ability to run more data efficiently. This would assist Steve for future data set and make it more efficient when he expands his data set to the entire Stock Market.

## Analysis and Challenges
### Analysis
Steve is attempting to find what stocks his parents should invest in based on the data from 12 companies 2017 and 2018 returns.
In order to assist Steve, I created a Worksheet that includes a table that finds the return of each stock per year. In this table we can see based on conditional formatting added which stocks have a positive rate of return and which stocks have a negative rate of return based on the color of the cell. Positive returns have a green cell and negative returns are colored red. 
In the original code, we nested the loops to find the total daily volume, starting prices, and endings prices, of each stock. In order to refactor the code I set the index to 0 before going further. Additionally, I created a stand alone loop to set the volume of each stock to zero. Together these factors created effieciencies before starting the loop that would find the total volume, starting prices, and ending prices for each stock because they did not have to run as many times.

## Summary
