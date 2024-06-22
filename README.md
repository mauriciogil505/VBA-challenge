# VBA-challenge
# This README file outlines the objectives and instructions to create a VBA script to analyze stock market data.

## Project Overview

This VBA project analyzes generated stock market data for quarters 1 through 4. The script is meant to extract key metrics.
We want to generate the following in a separate column in each sheet or tab within the Excel spreadsheet:
- Ticker Symbol
- Quarterly Change (from the opening price at the beginning of the quarter to the closing price at the end of the quarter)
- Percentage Change (from the opening price at the beginning of the quarter to the closing price at the end of the quarter)
- Total Stock Volume

The script also extracts and prints the ticker symbol with the:
- Greatest percentage increase
- Greatest percentage decrease
- Greatest total volume

## Instructions
Open Excel and navigate to "Developer" tab across the top of the menu optionps (this must be enabled and is required to complete this)
Then select "Visual Basic". It's located far left in the sub menu after selecting Developer (on Mac). 
After selecting Visual Basic, navigate to "insert" and select Module.
Now, you can create a VBA script in the new text editor window or "Module".

### Files Included for this Assignment

- Stock market data excel files are needed to have data to conduct data analysis.

### Creating the Script
First, start by declaring your variables
Declaring your variables allows for the VBA code to hold data such as ticker symbols, prices, volume, and calculations for the greatest increases and decreases.

Next, we want to use a "loop" function in order to iterate through each of the four quarters.

Then, we want to use conditional formatting to print results in green or red for positive or negative results for easier visualization.

This script will identify and print the greatest percentage increase, decrease, and total volume results into each sheet.

Credit: Stack Overflow, Class Recordings, Class Notes, Peer discussions, XPert Learning Assistant
