# The VBA of Wall Street

## Summary

This VBA script is used to analyze generated stock market data. 

### Files

* [Script](VbaChallenge.bas) - This script run through each worksheet and creates a summary table highlighting the change in stock price. 

* [ScreenShots](Capture-2018.PNG, Capture-2019.PNG, Capture-2020.PNG ) - Screen shot for each year of the summary table on the multi-year stock data


## Instructions

This script loops through all the stocks for one year and outputs the following information:

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

  * Script uses conditional formatting that will highlight positive change in green and negative change in red.

  * Script under the  **BONUS** section returns the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".

  * ```
    WS_Count = ActiveWorkbook.Worksheets.count
    For w = 1 To WS_Count
    ```
    is used to allow the VBA script to run on every worksheet (that is, every year) just by running the VBA script once.





