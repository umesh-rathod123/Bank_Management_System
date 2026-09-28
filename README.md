# Simple Banking System

A command-line based banking system developed using Python and Excel.

## Features

- Create Account
- Generate unique Account Number
- Deposit Money
- Withdraw Money
- Check Balance
- PIN verification
- Balance validation
- Insufficient balance checking
- Data storage using Excel
- Persistent data after closing the program

## Technologies Used

- Python
- Excel
- OpenPyXL

## Functions Used

- `create_account()`
- `deposit_money()`
- `withdraw_money()`
- `check_balance()`
- `find_account()`
- `verify_pin()`
- `main()`

## Excel Structure

The banking data is stored in:

`banking_system.xlsx`

Columns:

- Account Number
- Name
- Mobile Number
- PIN
- Balance

> Note: The Excel database is not included in this public repository because it may contain personal/account information.

## How to Run

Install OpenPyXL:

```bash
python -m pip install openpyxl
