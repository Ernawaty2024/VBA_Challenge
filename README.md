# VBA Challenge
## About The Project
This project comprises a VBA script designed tp systematically analyze stock data across various quarters. The script loops through all stocks for each quarter, extracting crucial information and generating insightful outputs:
1. **Ticker Symbol Extraction**: The script retrieves the ticker symbols for each stock.
2. **Quarterly Change Calculation**: It computes the quarterly change by subtracting the opening price at the beginning of a quarter from the closing price at the end of that quarter.
3. **Percentage Change Calculation**: Percentage change is detemined by comparing the opeing price at the start of a quarter with the closing price at the end of that quarter. 
4. **Total Stock Volume**: The script aggregates the total stock volume for each stock.
5. **Extreme Performer Identification**: The script is equipped the functionality to identify and return the stock exhibiting the "Greatest % increase", "Greatest % decrease", "and "Greatest total volume". This feature enables users to quickly pinpoint stocks with notable performance metrics.
6. **Enhance Script Flexibility**: To improbe usability and efficiency, the script has been modified to operate seamlessly across all worksheets, representing each quarter. This adjustment enables users to analyze data from multiple quarters with a single execution of the script, streamlining the analytical process.
7. **Visual Data Representation**: To facilitate intuitive data interpretation, the script incorporates conditional formatting. Positive changes are highlighted in green, while negative changes are highlighted in red. This visual cue assists users in identifying trends and patterns effortlessly.
## Usage
To utilize the script:
- Ensure the Excel workbook containing the <code style="color : blue">[Stock Data](Copy of Multiple_year_stock_data.xlsm)</code>  is open
- Navigate to the VBA editor
- Download the provided script file from <code style="color : blue">[VBA_Challenge Script_File](Copy of Multiple_year_stock_data.xlsm)</code>
- Paste the provided script into a new module
- Run the script to generate the desired outputs.
## Note
This script assumes that the stock data is organized in a structured format within an Excel workbook. Adjustments may be required based on the specific layout of the data.
## Credit
Data for this dataset was generated by edXBootCamps LLC and is intended for educational purposes only.
