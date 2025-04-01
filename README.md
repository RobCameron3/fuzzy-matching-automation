# Fuzzy Matching Automation

**Note:** This is a **demo version** of a Python-based fuzzy matching system originally developed in a professional setting. All data used here is sample/mock data and does not reflect any proprietary client information. The logic and workflow have been adapted to showcase the structure and methodology while maintaining data privacy.

This project automates account-matching tasks using Python and fuzzy logic to link unmapped records to existing account data. It includes built-in data cleaning, string standardization, and Excel export for manual validation and reporting.

## Overview

The script:
- Reads two Excel files: a CRM account list and a list of unmapped records.
- Standardizes fields like ZIP codes and addresses.
- Uses fuzzy matching to identify the top two most likely account matches.
- Outputs a structured Excel file with matches, scores, and account IDs.
- Prepares clean data for invoicing, analytics, or CRM updates.

## How It Works

1. Preprocessing: Cleans and formats address, ZIP, and account name data.
2. Fuzzy Matching: Applies fuzzy string logic (via `fuzzywuzzy`) to compare address fields within the same state or ZIP code.
3. Parallel Processing: Speeds up the match process using all available CPU cores.
4. Excel Export: Outputs a final file that includes original inputs, best match suggestions, and fuzzy match scores.

## Files

- `fuzzy_account_matcher.py`: Main Python script for data processing and matching.
- `data/sample_clients.xlsx`: Sample CRM account list.
- `data/unmapped_clients_sample.xlsx`: Sample of unmapped records to be matched.
- `data/client_report_by_zip.xlsx`: Final output with fuzzy match results.

## Technologies Used

- **Python**: pandas, numpy, fuzzywuzzy, joblib, datetime
- **Parallel Processing**: Accelerates performance for larger datasets
- **Excel Output**: Enables manual QA, invoicing, and final review

## Business Impact

This tool was used in a real-world setting to reduce manual matching time by 41%. It provided consistent results, supported developer collaboration, and improved downstream workflows by generating structured outputs for broader integration.

## How to Use

1. Replace the sample Excel files with your own in the `/data` folder.
2. Run the script:

3. Open `client_report_by_zip.xlsx` in the `/data` folder to review the results.
