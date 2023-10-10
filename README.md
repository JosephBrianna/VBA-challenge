# Stock Analysis VBA Script

## Overview

This VBA script is designed to analyze a list of stock data, calculate various metrics for each stock, and summarize the results. It loops through the data and outputs the following information for each stock:

1. Ticker symbol
2. Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
3. The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
4. Total stock volume of the stock.

## Usage

1. **Open Excel:** Make sure you have Microsoft Excel open and the workbook containing your stock data loaded.

2. **Worksheet Setup:**
   - Ensure that your stock data is organized with columns in the following order:
     - Column A: Ticker symbol
     - Column C: Opening price at the beginning of the year
     - Column F: Closing price at the end of the year
     - Column G: Total stock volume
   - The data should be sorted by ticker symbol and date in ascending order.

3. **Insert VBA Script:**
   - Press `ALT` + `F11` to open the VBA editor in Excel.
   - In the VBA editor, click `Insert` > `Module` to insert a new module.
   - Copy and paste the provided VBA script into the module.

4. **Run the Script:**
   - Close the VBA editor.
   - In Excel, press `ALT` + `F8` to open the "Macro" dialog box.
   - Select "StockAnalysis" from the list and click "Run."

5. **View Results:**
   - After running the script, a new summary table will appear in your worksheet.
   - This table will contain the analyzed data with ticker symbols, yearly changes, percentage changes, and total stock volumes.

6. **Format Percentage Column (Optional):**
   - The percentage change column (column K) may not be formatted as percentages initially. You can format this column as a percentage by selecting the cells (K2 to the last row with data) and applying a percentage number format.

## Important Notes

- Ensure your stock data is correctly formatted and sorted before running the script.
- Make a backup of your data before running any VBA script to prevent data loss in case of errors.
- The script assumes that the stock data is in the same workbook where you're running the script.

## Disclaimer

This script is provided as a tool for analyzing stock data and should be used responsibly. It may require adjustments to match your specific data format and layout. Use caution and verify the results to ensure accuracy.

---

Feel free to customize this README file to include any additional instructions or notes specific to your use case.
