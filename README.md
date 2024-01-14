# VBA-challenge


# Stock Market Summary Excel VBA Script

This Excel VBA script is designed to analyze and summarize stock market data across multiple worksheets in an Excel workbook. The script calculates and presents key metrics for each stock ticker, including yearly change, percent change, and total stock volume. Additionally, it generates a secondary summary table highlighting stocks with the greatest percent increase, percent decrease, and total volume across all sheets.

## Features

- Dynamic Summarization: The script dynamically analyzes each worksheet, identifying unique stock tickers and calculating their yearly change, percent change, and total stock volume.

- Conditional Formatting: The script utilizes conditional formatting to visually highlight positive and negative yearly changes, enhancing data visualization.

- Secondary Summary Table: A separate summary table is generated, showcasing the stocks with the greatest percent increase, percent decrease, and total volume across all worksheets.

## Usage

1. Add Data:
    - Populate each worksheet with stock market data, ensuring that each row represents a unique day of trading for a particular stock.

2. Run the Script:
    - Open the Excel workbook containing the data.
    - Press `ALT + F8` to open the "Macro" dialog.
    - Select `StockMarket_summary` from the list and click "Run."

3. Review Results:
    - The script will dynamically generate summary tables on each worksheet, providing insights into individual stock performance and overall market trends.
