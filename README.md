## Mod 2 VBA
 

# Overview

This VBA script processes stock market data across all worksheets in an Excel workbook. It extracts stock tickers, calculates total volume, quarterly change, and percentage change, and outputs the results in a summary table with conditional formatting.

# Features

Collect Unique Tickers

Identifies and gathers all unique stock tickers present in the dataset.

Calculate Key Metrics

Computes the total volume for each stock ticker.

Determines the price difference between the opening and closing values.

Calculates the percentage change in stock price.

# Apply Conditional Formatting

Highlights the price difference:

Green for a positive change.

Red for a negative change.

Applies similar formatting for percentage change.

# Additional Functionality

The script identifies and returns the stock with:

Greatest % Increase

Greatest % Decrease

Greatest Total Volume

The output format is designed to match the expected structured summary table.