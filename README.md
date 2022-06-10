# VBA-challenge: The VBA of Wall Street

This repository contains a VBA script file called runAll which runs through all other VBA scripts in the repository; vba_Challenge and then bonus_Challenge. The script can be ran via a button within the workbook Multiple_year_stock_data.

VBA Scripts:
[runAll_Script](runAll_Script.vbs)
[vba_Challenge_Script](vba_Challenge_Script.vbs)
[bnus_Challenge_Script](bonus_Challenge_Script.vbs)

The vba_Challenge script creates 4 columns on each worksheet within the workbook: Ticker, Yearly Change, Percent Change, and Total Stock Volume. The columns contain the following:

  * All unique ticker symbols from the raw data.

  * The yearly change from opening price at the beginning of a given year to the closing price at the end of that year, for each ticker. The cells in this column are formatted based on value; green for a positive value, red for a negative value, and grey if the value is 0.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year, for each ticker.

  * The total stock volume of the stock, for each ticker.

The bonus_Challenge script creates a summary table on each worksheet within the workbook. It identifies the value and the respective ticker that observed:

  * The greatest percent increase.

  * The greatest percent decrease.

  * The greatest total stock volume. 


Screen shots of yearly results:

[2018_Results](2018_Results.pdf)
[2019_Results](2019_Results.pdf)
[2020_Results](2020_Results.pdf)