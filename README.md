# Report and Priority Cases Analysis

## Overview

This script automates the process of fetching sales report files from Google Drive, extracting data from Google Sheets, and updating a priority cases sheet with analyzed data. It performs the following tasks:
1. Authenticates with Google Drive and Google Sheets APIs.
2. Fetches sales report files created in the current month.
3. Extracts data from specified Google Sheets.
4. Updates the priority cases sheet with data from the sales reports.

## Requirements

- Python 3.x
- Required Python libraries:
  - os
  - pickle
  - pandas
  - datetime
  - google-auth
  - google-auth-oauthlib
  - googleapiclient
  - gspread
  - json
  - re

## Setup and Installation

1. Clone the repository or download the script file.
2. Install the required libraries using pip:
   pip install pandas google-auth google-auth-oauthlib googleapiclient gspread
3. Ensure you have `credentials.json` for Google Drive and Google Sheets API credentials and place it in the same directory as the script.

## Authentication with Google APIs

The function `authenticate()` handles the authentication process. It checks if a token exists and is valid. If not, it initiates the OAuth flow to get a new token.

### authenticate()
- **Purpose**: Authenticates with Google Drive and Google Sheets APIs using OAuth 2.0.
- **Process**:
  1. Checks if `token.json` exists to load credentials.
  2. If credentials are expired or do not exist, initiates the OAuth flow using `credentials.json`.
  3. Saves the new credentials in `token.json`.
- **Returns**: Google credentials object.

## Fetching Sales Report Files

### get_sales_report_files()
- **Purpose**: Fetches sales report files created in the current month from Google Drive.
- **Process**:
  1. Defines the query to search for sales report files created in the current month.
  2. Calls the Drive API to list files matching the query.
  3. Extracts and returns the URLs of the found files.
- **Returns**: List of URLs of sales report files.

## Extracting Data from Google Sheets

### get_sheet_data(sheet_url, sheet_name)
- **Purpose**: Extracts data from a specified Google Sheet.
- **Parameters**:
  - `sheet_url`: URL of the Google Sheet.
  - `sheet_name`: Name of the sheet to extract data from.
- **Process**:
  1. Opens the Google Sheet by URL.
  2. Extracts all records from the specified sheet.
  3. Converts the records to a pandas DataFrame.
- **Returns**: pandas DataFrame with extracted data.

## Reading Priority Cases and Mastersheet

### Reading Priority Cases Sheet
- **Purpose**: Reads data from the priority cases sheet.
- **Process**:
  1. Opens the priority cases sheet by URL.
  2. Extracts all records from the 'KAM' sheet.
  3. Converts the records to a pandas DataFrame.
- **Returns**: pandas DataFrame with priority cases data.

### Reading Mastersheet
- **Purpose**: Reads data from the mastersheet.
- **Process**:
  1. Opens the mastersheet by URL.
  2. Extracts all records from the 'Mastersheet' sheet.
  3. Converts the records to a pandas DataFrame.
- **Returns**: pandas DataFrame with mastersheet data.

## Updating Priority Cases Sheet

### update_priority_cases(priority_cases_df, daily_report_df, date_str)
- **Purpose**: Updates the priority cases sheet with data from daily report sheets.
- **Parameters**:
  - `priority_cases_df`: DataFrame with priority cases data.
  - `daily_report_df`: DataFrame with daily report data.
  - `date_str`: Date string extracted from the daily report.
- **Process**:
  1. Matches student names between priority cases and daily report data.
  2. Updates the priority cases DataFrame with message counts for the matched students.
  3. Prints debug information for matched and unmatched cases.
- **Returns**: None.

## Processing Daily Reports

### Processing Each Daily Report Sheet
- **Purpose**: Processes each daily report sheet to update the priority cases DataFrame.
- **Process**:
  1. Iterates over the list of daily report URLs.
  2. Extracts data from each daily report sheet.
  3. Updates the priority cases DataFrame with data from the daily report.
  4. Sorts the date columns and calculates total and count of message counts.
  5. Reorganizes and updates the priority cases sheet with the new data.
- **Returns**: None.

## Running the Script

To run the script, execute it in your Python environment. The script will authenticate, fetch sales report files, extract and process data, and update the priority cases sheet.

## Conclusion

This script automates the process of fetching, extracting, and updating data from Google Drive and Google Sheets. Follow the setup instructions to get started and refer to the function descriptions for detailed information on each component of the script.
