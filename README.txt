# Top Transactions Per Month

This application processes a fraud analysis Excel file and generates monthly reports showing the top cards and cashiers with the most transactions.  
It saves the results into an output folder called **TopTransactionsPerMonth**.

---

## Features
- Reads input Excel file (e.g., `fraud_analysis.xlsx`).
- Processes transactions by **month**.
- Creates an Excel file for each month containing:
  - **RawData**: All transactions for that month (without extra processing columns).
  - **TopCards**: Summary of the top N cards with the most transactions, including:
    - Card Number  
    - Month  
    - Total Transactions  
    - First & Last Transaction  
    - Distinct Branches, Cashiers, and Registers  
    - Day with Most/Fewest Transactions  
    - (Optional) Sum of `trans_total` if column exists
  - **TopCashiers**: Summary of the top N cashiers with the most transactions, including:
    - Cashier ID  
    - Month  
    - Total Transactions  
    - First & Last Transaction  
    - Distinct Branches, Cards Handled, and Registers  
    - Day with Most/Fewest Transactions  
    - (Optional) Sum of `trans_total` if column exists

---

## Requirements
- Python 3.11+
- Dependencies:
  - `pandas`
  - `openpyxl`

---

## How to Run 
-- Step 1:
    - Open CMD and go to Directory /VScan

-- Step 2:
    - virtualenv venv 

-- Step 3: 
    - pip install -r requirements.txt

-- Step 4: 
    - python app.py 

-- Step 5:
    - Upload .csv/.xlsx

-- Step 6: 
    - Input the Number of Top Cards and Top Cashiers to be Generated




