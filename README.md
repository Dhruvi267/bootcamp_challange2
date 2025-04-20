# 📈 VBA Stock Market Analysis

## Overview

This project analyzes stock market data using **VBA scripting in Excel**. It automates the process of calculating important financial indicators like **quarterly change**, **percentage change**, and **total trading volume** for each stock ticker across multiple sheets representing different quarters or categories. The project is part of a VBA scripting challenge designed to demonstrate automation, logic building, and data analysis skills in Excel.

---
## Features

✅ **Automates Analysis** of large stock datasets  
✅ **Calculates**:
- Total Volume per ticker  
- Quarterly Change (Close - Open)  
- Percentage Change ((Close - Open) / Open)  

✅ **Highlights**:
- Green for positive change  
- Red for negative change  

✅ **Identifies**:
- Greatest % Increase
- Greatest % Decrease
- Greatest Total Volume  

✅ **Runs on All Sheets**: No manual repetition required

---

## How It Works

The script loops through each worksheet (A–F or Q1–Q4), analyzes the stock data row by row, and outputs the results in a clean summary format with conditional formatting and performance metrics.

### Output Columns:
| Column | Description |
|--------|-------------|
| `Ticker` | Stock symbol |
| `Quarterly Change` | Difference between closing and opening prices |
| `Percent Change` | Change as a percentage of the opening price |
| `Total Volume` | Sum of stock volume for that ticker |

### Summary Metrics (per sheet):
| Metric | Meaning |
|--------|---------|
| Greatest % Increase | Highest growth by percentage |
| Greatest % Decrease | Largest drop by percentage |
| Greatest Volume | Ticker with the highest trading volume |

---

## How to Run the Script

1. Open Excel and load either `alphabetical_testing.xlsx` or `Multiple_year_stock_data.xlsx`
2. Press `Alt + F11` to open the VBA Editor
3. Insert a new **Module** (Right-click ➝ Insert ➝ Module)
4. Paste the appropriate script
   - `AnalyzeAlphabeticalStocks` for `alphabetical_testing.xlsx`
   - `AnalyzeQuarterlyStockData` for `Multiple_year_stock_data.xlsx`
5. Press `F5` to run the script

---

## Sample Results

![results](https://github.com/user-attachments/assets/b62e085d-06d9-49d1-a501-1800f47fa390)


---
## Skills Demonstrated

- Visual Basic for Applications (VBA)
- Automation in Excel
- Data Cleaning & Analysis
- Conditional Formatting
- Multi-sheet Logic & Looping
- Financial Metric Computation
- Git & Version Control

---

