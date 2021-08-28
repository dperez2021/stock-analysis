#Overview of Project
------------------
## Purpose

The purpose of this project is to complete an analysis of green energy stocks using VBA for Steve so that it could be presented to his parents. This project will measure performance over the years of 2017 & 2018. With the data given we will be able to measure the trading volumes and total returns. The project will conclude with a touch of a button to analyze our data.

##Background
The data that is provided for analyzing includes 12 different stocks with two years of data which include; openning, closing, adjusted closing price, highest price, lowest price, and the volume of the stock. 

#Results
----
##Analysis

Once the script is run we are able to compare the 2017 and 2018 results for the 12 green energy stocks that were isolated. The returns during 2017 favored much better then 2018 based on the percentage of return. Volume was for the most part consistant over 3,000,000,000 stocks traded throughout each year. In 2018 only 2 stocks ENPH and RUN traded for a positive return, whereas in 2017 only 1 stock traded for a negative return TERP.

## Comparing the Original Script vs the Refractored Script.

IMAGE?

The run time for the original code when inputing 2017, took 1.16 seconds.

IMAGE

The run time for the original code when inputing 2018, took 1.15 seconds.

IMAGE

Below is the original code before refractoring.

IMAGE

The run time for the refractored code when inputing 2017, took 0.20 seconds.

IMAGE

The run time for the refractored code when inputing 2018, took 0.20 seconds.

IMAGE

Below is the refractored code.

IMAGE

Refactoring the code reduced the run time by approximately 1 second. The difference was that we inserted an array which allowed me to assign variables of tickerVolumes, tickerStartingPrices, and tickerEndingPrices to each ticker symbol before interating through the data set. By doing it this way the code would run faster because we are not using nested for loops like in the original code.

#Summary of Refactoring
----
##Advantage and Disadvantage

The advantages of refactoring is that the run time should decrease, and the code may shrink a bit as you consolidate and revise the code.

The disadvantage is while review the code to refactor it you may create unusable code which may delay the time it takes to successfully refactor the code.

##Refactoring VBA Script
The pro's of refactoring is that you already have a working code and you can place this code side by side and tweak it line by line. As you tweak the code you can run to see what difference it makes. The con's of refactoring is that changing one letter can make the entire code unusable. The knowledge of understanding syntax is very important when trying to make the original code more efficient.

