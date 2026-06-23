# Access Reconciliation Tool

## Overview
An automated data reconciliation tool that compares cardholder records between two physical security systems — C-CURE and Genetec — and identifies discrepancies.

## Problem It Solves
Security teams manually compared thousands of records between two access control systems every week. This took hours and was prone to human error. This tool automates the entire process in seconds.

## How It Works
1. Reads two Excel files from the data/ folder
2. Compares card numbers across both systems
3. Generates a timestamped Excel report with 3 tabs:
   - Only in C-CURE
   - Only in Genetec
   - In Both

## Technologies Used
- Python 3.10
- pandas
- openpyxl
- Git/GitHub

## How To Run
pip install -r requirements.txt
python3 reconcile.py

## Author
Anointed-David
