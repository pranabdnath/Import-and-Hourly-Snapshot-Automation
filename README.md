# Import-and-Hourly-Snapshot-Automation
This script automates two tasks inside Google Sheets:  Importing data from an external source sheet.  Creating an hourly snapshot that summarises city-wise counts and threshold-based counts.  It helps maintain an updated view of data while also generating hourly summaries without manual work.


Problem Statement

Working with large data sheets required frequent importing from another sheet and calculating city-wise summaries. Doing this manually was time-consuming and inconsistent.
There was no automated way to:

Pull the latest data at fixed intervals

Generate summaries based on changing values

Maintain a clean and updated structure inside the working sheet

The goal was to automate the entire workflow using Google Apps Script.

What This Script Does

Imports the complete dataset from another spreadsheet.

Replaces existing data in the target sheet with the latest values.

Reads selected columns to extract information such as:

City names

Count per city

How many rows meet a threshold value

Writes the computed summary into dynamically decided columns.

Resets the summary area before writing new values.

How the Logic Works

The script fetches city names and numeric values.

It identifies unique cities.

For each city, it calculates:

Total row count

Count of rows where a specific value is equal to or above 15

The output is written to the sheet based on the current hour.

Morning data is placed in one column set; other hours use another.
