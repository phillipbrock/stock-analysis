# Stock Analysis
## Overview of Project
The purpose of this coding project and subsequent analysis is to allow a financial professional to present data to their clients for investment decisions. The specific analysis, in this case, measures and compares stock performance in respective years for certain green energy companies as needed by the client. The basic coding language will allow the financial professional to automate and apply this analysis to data involving different companies or years as needed.
## Results
In order to run this analysis, I created two VBA macros which automate the analysis. Both macros perform the same function, though the latter has been refactored to run more efficiently. The anaylsis calculates the total daily volume for each stock in each respective year, as well as the return for each respective year.

**2017**

![2017](https://github.com/phillipbrock/stock-analysis/blob/main/Resources/Stock_Performance_2017.PNG)

**2018**

![2018](https://github.com/phillipbrock/stock-analysis/blob/main/Resources/Stock_Performance_2018.PNG)

In order to make the analysis easier more visually intuitive, I formatted the return column so that cells representing positive returns are green and those representing negative returns are red using this VBA code in the macro:

![formatting code](https://github.com/phillipbrock/stock-analysis/blob/main/Resources/code%201.PNG)

I created two different macros for this analysis, one of which is lengthier and runs less efficiently, then a refactored version which runs more efficiently but with the same results for the analysis. The efficiency differential can be shown by including the run time in the macro coding.

**Using Original Code: 2017**

![unrefactored](https://github.com/phillipbrock/stock-analysis/blob/main/Resources/VBA_unrefactored_2017.PNG)

**Using Refactored Code: 2017**

![refactored](https://github.com/phillipbrock/stock-analysis/blob/main/Resources/refactored_2017.PNG)

The results show 2017 to be an overall good year for the selected stocks with two exceptions. 2018 was the opposite, with negative returns for all but two of the stocks. The stock with ticker callsign "TERP" saw negative returns both years and presumably is to be avoided. Stocks "RUN" and "ENPH" saw positive returns both years, with the latter having an extraordinary return in both years, as well as a high daily volume. More thorough insights can likely be gleaned by the financial professsional.
