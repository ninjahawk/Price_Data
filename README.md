# Price_Data

This program automates the **Price Data Sheet** by pulling live price data from Yahoo Finance and updating the Excel file.

---

## Features
- Pulls live stock data for the **"DROP DATA HERE"** tab.
- Overwrites existing data with updated values.
- Fetches the following metrics:
  - **Last**: Last available stock price (close).
  - **Change**: Difference between the open price and the current price.
  - **% Change**: Percentage change relative to the open price.
  - **Open**: Opening price for the day.
  - **High**: Highest price for the day.
  - **Low**: Lowest price for the day.
  - **Volume**: Trading volume for the day.
  - **Time**: Date the data was fetched.

---

## Current Excel File Path

---

## Prerequisites
Make sure you have the following installed:
- **Python 3.x**
- **yfinance** (for fetching stock data)
- **openpyxl** (for editing Excel files)

Install dependencies using `pip`:
```bash
pip install yfinance openpyxl