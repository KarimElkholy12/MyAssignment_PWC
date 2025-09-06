# RPA Automation - Data Cleansing , consolidations and Web scraping

## Task Overview
This project analyzes employee data from an Excel file (`Employee Raw.xlsx`) and generates 2 reports with key metrics.
1. Employees_Cleaned.xlsx
2. Employees_Splitted.xlsx

Additionally, web scraping from finance yahoo portal and retrieve data.


## Steps Performed
The automation performs the following steps:

1. Remove duplicate rows
2. Standardize dates into YYYY-MM-DD format.
3. Format phone numbers into +20XXXXXXXXXX.
4. Trim spaces in names.
5. split emails to two groups.
6. Web scraping - Extract the Stock Symbol, Name, Price, Change, % Change, and Volume from the Most Active Stocks table.


## Edge Cases Handled

The script robustly handles the following edge cases:

1. **Input file not found**
   - Displays warn log message: "Error: No File found"

## Assumptions
Assumed that below data can be modified later, therefore placed in Config file for better maintenance and reusability.
1. File paths.
2. Required date format.
3. Finance Yahoo portal url.

## Time taken
4 hours

# Employee Data Analysis Python Script

## Task Overview
This project analyzes employee data from an Excel file (`Employees_Cleaned.xalsx`) and generates a summary report with key metrics.

## Analysis Performed
The script performs the following analysis:
1. **Count unique companies** - Identifies the total number of distinct companies
2. **Top 5 email domains** - Finds the most common email domains among employees
3. **Employee distribution** - Counts the number of employees per company

## Solution Approach

### Data Processing Steps
1. **Read Excel file** - Load `Employees_Cleaned.xlsx` using pandas
2. **Extract email domains** - Parse email addresses using regex to extract domain names
3. **Aggregate data** - Count unique values and group data by company
4. **Generate report** - Format results and save to CSV

### Key Implementation Details
- Used pandas for efficient data manipulation
- Applied regex pattern `@([^\s]+)` to extract email domains
- Used value_counts() for frequency analysis
- Structured output in a clear, readable CSV format

## Requirements
- Python 3.x
- pandas library
- openpyxl library (for Excel file reading)

## Installation Instructions

Using pip3
```bash
pip3 install pandas openpyxl
```

## How to Run

1. **Ensure the Excel file is present**
   - Place `Employees_Cleaned.xlsx` in the same directory as the script

2. **Run the analysis script**
   ```bash
   python employee_analysis.py
   ```
   Or if using python3:
   ```bash
   python3 employee_analysis.py
   ```

3. **Check the output**
   - Console will display the analysis results
   - `Summary_Report.csv` will be created with detailed results

## Output Files
- **Summary_Report.csv** - Contains all analysis results in structured format:
  - Total unique companies count
  - Top 5 email domains with employee counts
  - Employee count per company

## File Structure
```
.
├── Employees_Cleaned.xlsx  # Input Excel file
├── employee_analysis.py    # Analysis script
├── Summary_Report.csv      # Output report
└── README.md              # This file
```

## Edge Cases Handled

The script robustly handles the following edge cases:

1. **Empty Excel file**
   - Displays error message: "Error: The Excel file contains no data"
   - Exits gracefully

2. **Missing required columns**
   - Checks for 'Company' and 'Email' columns
   - Shows which columns are missing
   - Exits with clear error message

3. **File not found**
   - Displays: "Error: Employees_Cleaned.xlsx not found in current directory"
   - Exits before attempting to read

4. **Less than 5 email domains**
   - Shows all available domains (e.g., if only 4 exist)
   - Adds note: "only X email domains found" for remaining slots
   - Maintains consistent report structure

5. **No email domains found**
   - Handles cases where all emails are invalid or missing
   - Reports: "No email domains found" in all 5 slots

6. **Invalid email formats**
   - Safely extracts domains using regex
   - Returns None for invalid/missing emails
   - Continues processing valid entries

7. **Null/NaN company names**
   - Filters out invalid company names
   - Only counts valid company entries

8. **No companies with employees**
   - Displays: "No companies found with employees"
   - Still generates report with appropriate message

9. **Excel read errors**
   - Catches exceptions during file reading
   - Displays specific error message
   - Exits gracefully


## Sample Output
```
Analysis Complete!
Number of unique companies: 6

Top 5 email domains:
  corp.com: 8 employees
  yahoo.com: 8 employees
  gmail.com: 5 employees
  company.com: 4 employees

Employees per company:
  Alpha Ltd: 2 employees
  Beta Corp: 3 employees
  Delta LLC: 6 employees
  Gamma Inc: 3 employees
  Internal: 8 employees
  Omega Group: 3 employees

Results saved to Summary_Report.csv
```
## Time taken
10 hours
