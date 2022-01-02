# stock-analysis

## Overview of Project 

Our client Steve has asked us to prepare him a VBA workbook that will allow him to analyze various stocks over the past decade for his parents' investment pursuits. For the purpose of this project, we have refactored a new script from code used originally to analyze a dozen stocks from 2017 and 2018. For Steve, it would be beneficial for him to have a script that allows him to analyze thousands of stock over several years if he wishes. The original dataset we used to develop our code was green_stocks.xlsm and the programming language we used was Excel VBA.

## Results

After running our code, we compared the performance of 12 green stocks between 2017 and 2018. In our header, "Ticker" describes the specific stock (abbreviated), "Total Daily Volume" describes the average number of stocks traded daily, and "Return" describes the percentage increase or decrease in price from the beginning of the year to the end of the year. The stock analysis we performed shows that:

* Eleven out of the twelve stocks in 2017 had positive yearly returns. 

  * Only one stock, TERP, had a negative yearly return of "-7.2%".
  * The execution time of the original script for 2017 is approximately 0.76 seconds.
  * The execution time of the refactored script for 2017 is approximately 0.16 seconds.

<p float="left">
<img src="https://github.com/katmarcin/stock-analysis/blob/b5b717138e9d927d1eabdaab105e8bdf0fcc072b/2017_Data.png" width="245" height="245" />
<img src="https://github.com/katmarcin/stock-analysis/blob/0561102ccf086b6aa594ef5741ec0f8a09d57ce3/2017_Original_Runtime.png" width="275" height="245" />
<img src="https://github.com/katmarcin/stock-analysis/blob/32428f5f451d661ed2f4fcfe7aedefe1b89b1d8c/Resources/VBA_Challenge_2017.png" width="275" height="245" />
</p>

* Ten out of twelve stocks in 2018 had negative yearly returns.

  * Only two stocks, ENPH and RUN, had positive yearly returns. ENPH had a return of "81.9%" and RUN had a return of "84.0%".
  * The execution time of the original script for 2018 is approximately 0.83 seconds.
  * The execution time of the refactored script for 2018 is approximately 0.24 seconds.


<p float="left">
<img src="https://github.com/katmarcin/stock-analysis/blob/b5b717138e9d927d1eabdaab105e8bdf0fcc072b/2018_Data.png" width="245" height="245" />
<img src="https://github.com/katmarcin/stock-analysis/blob/0561102ccf086b6aa594ef5741ec0f8a09d57ce3/2018_Original_Runtime.png" width="275" height="245" />
<img src="https://github.com/katmarcin/stock-analysis/blob/32428f5f451d661ed2f4fcfe7aedefe1b89b1d8c/Resources/VBA_Challenge_2018.png" width="275" height="245" />
</p>


If Steve had to choose one stock from this dataset for his parents' portfolio, the best option would be "RUN". In 2017, "RUN" had a "5.5%" return and in 2018, that figure increased astronomically to "84.0%". This change indicates that this particular business is doing well and that the stock could continue to deliver as a long-term investment. 


The major change to the original VBA script that was made to obtain a more efficient, refactored code included rearranging our output arrays. With our refactored output arrays, our machine can skip through a ticker it has already read instead of having to re-read throught the data, allowing us to save time, productivity, and resources. In other words, once information for one stock is completely analyzed, the next stock can be read without having to go through the previous stock first. By creating the "tickerIndex" variable, we were able to assign our other arrays, such as "tickerVolumes" or "tickerEndingPrices", to the tickers array. Compared to the nested for-loop method used in the original code, which requires our machine to loop through our tickers from start to finish repeatedly, the refactored code proves to be more efficient.

Refactored code: 

<img src="https://github.com/katmarcin/stock-analysis/blob/4d991eed4f1c3fa3e0313ff0bebd549c120a0daf/Output_arrays.png" width="340" height="400">

Original code: 

<img src="https://github.com/katmarcin/stock-analysis/blob/6c9d7ca9109f2915ef93ab4fe3e1ea0cc8c997eb/Original%20code.png" width="340" height="400">


## Summary

  Generally, refactoring code may have more advantages than disadvantages. However, this is all subject to the particular dataset and the information that is desired. As visualized in the results, refactored scripts were generated about half a second faster for both 2017 and 2018 than both years in the original VBA script. As a reminder, this dataset is relatively short.  Efficiency is a major advantage, especially for our client Steve, as he wishes to apply this code for thousands of stocks. This difference could equate to several seconds. The reason why this code runs faster, is due to another advantage, improved code logic. When complexity is tweaked and reduced, less steps can be taken which in turns uses less memory. A disadvantage to refactoring code is the difficulty in locating fallacies in code logic when an error is produced. Debugging may not always provide a clear solution to a code that is not running properly. For example, a challenge we encountered in our refactoring of our code was a persistent "Next without For" compile error. After several minutes of troubleshooting, we discovered a missing "End If" in our code. It is important to ensure for-loops are properly edited to prevent errors either from missing information, improper indentation, etc.  This leads us to our next disadvantage, loss of time or productivity. When refactoring, one must consider the possibility of code breaking and setting aside time for troubleshooting. If one is careful and meticulous, refactoring code proves to be more advantageous.
