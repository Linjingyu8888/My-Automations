# My-Automations
This Python code automates the task of consolidating data from multiple Excel files (.xls or .xlsx) located in a specified directory into a single CSV file. Here is a detailed explanation of the problem it solves:

## Problem Being Solved
Data Aggregation Across Files:

You have a directory containing multiple Excel files.
Each file has data in the first sheet that you want to extract and merge into one consolidated file.
Uniform Formatting:

The output data needs to be saved as a single CSV file with all rows from all input Excel files.
Handling Headers and Skipping Rows:

If the Excel files contain headers (or unwanted rows) at the top, these can be skipped using the number_of_rows_to_skip parameter.
Adding Metadata:

Each row in the final consolidated CSV file includes the file name from which the data originated for traceability.
Error Handling and File Filtering:

Ignores non-Excel files and user-specified files in the directory.
Ensures only .xls and .xlsx files are processed.

## Key Features
### Dynamic Headers:
Extracts the header row dynamically from each Excel file, starting from a specified row (useful for inconsistent Excel formats).
### Date Formatting:
Converts datetime values in Excel files into ISO 8601 date strings for consistent formatting.
### Logging and Feedback:
Provides feedback on processed files and skipped files directly in the console.

## How It Works
1. Input Directory Scanning:

The script scans the given directory for all files, filtering out non-Excel files and files explicitly listed in files_to_ignore.

2. File Reading:

For .xls files: Uses xlrd to read the data.
For .xlsx files: Uses openpyxl.

3. Row Processing:

Skips the specified number of rows in each file.
Extracts the header row from the first data file and appends the data rows from all subsequent files.

4. CSV Writing:

Writes all rows to a single CSV file, including the file name for traceability.

5. Error Handling:

Catches and reports any issues, providing a traceback for debugging.