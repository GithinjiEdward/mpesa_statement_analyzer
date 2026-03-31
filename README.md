# M-PESA Financial Analyzer 

A Streamlit-based financial analysis app for password-protected M-PESA statements.

## Overview
This project helps users:
- Upload and unlock password-protected M-PESA PDF statements
- Extract and classify transactions
- Analyze money in, money out, balances, and transaction patterns
- Detect recurring payments
- Visualize transaction trends
- Simulate hire purchase affordability
- Optimize loan affordability based on cash flow and expense reduction options
- Export filtered analysis to Excel

## Features
- PDF statement upload and password unlock
- Transaction parsing and categorization
- Overview dashboard with charts and summaries
- Transaction-level filtering
- Recurring payment analysis
- Hire Purchase Simulator
- Loan Optimizer
- Excel export of filtered analysis

## Tech Stack
- Python
- Streamlit
- Pandas
- Plotly
- pdfplumber
- pikepdf
- XlsxWriter

## Project Structure
```text
mpesa_statement_analyzer/
├── app.py
├── requirements.txt
├── README.md
└── .gitignore

## Installation
git clone https://github.com/GithinjiEdward/mpesa_statement_analyzer.git
cd mpesa_statement_analyzer
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py

## Usage
1. Launch the app
2. Upload your M-PESA PDF statement
3. Enter the PDF password
4. Process the statement
5. Explore the dashboard, recurring payments, hire purchase simulation, and loan optimization.

## Deployment
This app can be deployed on Streamlit Community Cloud using the GitHub repository.

##Author
Edward Githinji

Github: https://github.com/GithinjiEdward
