# Excel to Financial Statements

A Python-based tool to convert Excel journal entries into detailed financial reports, including **Balance Sheet** and **Profit & Loss (P&L)** statements. This script reads journal entries from an Excel file, categorizes them into financial categories, and generates formatted reports in an Excel output file. Ideal for accountants and finance professionals looking to automate financial report generation.

## Features

- Convert journal entries in Excel into Balance Sheet and P&L statements.
- Automatically categorize entries into **Assets**, **Liabilities**, **Equity**, **Income**, and **Expenses**.
- Generate well-formatted Excel reports.
- Easy-to-use and customizable for different business needs.

## Requirements

Ensure you have Python and the following packages installed:

```bash
pip install pandas openpyxl
```

## Input File Format

The input Excel file (`journal_entries.xlsx`) should have the following columns:

| Date       | Description     | Type       | Debit | Credit |
|------------|-----------------|------------|-------|--------|
| 2024-01-01 | Cash            | Asset      | 1000  | 0      |
| 2024-01-02 | Accounts Payable | Liability  | 0     | 500    |
| 2024-01-03 | Revenue         | Income     | 0     | 1500   |
| 2024-01-04 | Rent            | Expense    | 300   | 0      |

- **Date**: Date of the transaction.
- **Description**: Description of the transaction.
- **Type**: Transaction type – acceptable values are `Asset`, `Liability`, `Equity`, `Income`, `Expense`.
- **Debit/Credit**: The debit and credit amounts for each entry.

## Usage

### 1. Clone the Repository

```bash
git clone https://github.com/dinuth-perera/excel-to-financial-statements.git
cd excel-to-financial-statements
```

### 2. Prepare Your Input File

Ensure you have an Excel file named `journal_entries.xlsx` with the correct format as described above.

### 3. Run the Script

```bash
python main.py
```

### 4. Generated Output

The script will generate a new Excel file named `financial_reports.xlsx` in the same directory. The file contains:
- **Balance Sheet**
- **Profit & Loss Statement**

#### Example Output

**Balance Sheet Example:**

| Category    | Amount  |
|-------------|---------|
| Assets      | 1000    |
| Liabilities | 500     |
| Equity      | 500     |
| **Total Equity and Liabilities** | **1000** |

**Profit & Loss Example:**

| Category    | Amount  |
|-------------|---------|
| Income      | 1500    |
| Expenses    | 300     |
| **Profit/Loss** | **1200** |
