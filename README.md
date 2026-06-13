# HCS Billing Automation

## Problem Statement
In a custom cloud zone where prices are hidden from enterprise end users, monthly billing must be produced from exported reports. HCS MOne can export daily usage sheets for each resource and user. This script assigns the correct prices to cloud services and computes the monthly total bill.

This tool is designed to work with Excel reports exported from **HCS MOne** and is intended for internal use by Huawei infrastructure administrators.

## Features
- Prompts the user for the path to the billing Excel file.
- Processes the exported report using `openpyxl`.
- Summarizes monthly net usage per resource per user.
- Generates a new Excel file (`result.xlsx`) with calculated billing data.

## Requirements
- Python 3.9+
- openpyxl

## Notes:
- The script is a lightweight internal automation tool, not a full billing system.
- The script expects the HCS MOne export format; other formats may fail.
