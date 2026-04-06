# Excel-Python Calculator using xlwings

## Overview
This project demonstrates Excel ↔ Python integration using xlwings.  
It allows users to perform mathematical operations in Excel while leveraging Python for computation and error handling.

---

## Features
- Excel-based UI for input/output
- Python backend for calculations
- Integration using xlwings
- Robust error handling
- Supports multiple operations:
  - SUM
  - PRODUCT
  - SUBTRACT
  - DIVIDE
  - POWER

---

## Project Structure

excel-python-calculator/
│── calculator.py        # Python logic + xlwings entry point
│── Calculator.xlsx      # Excel interface with inputs, outputs, and Compute button
│── README.md            # Project documentation
│── requirements.txt     # Python dependencies

---

## Setup Instructions

### 1. Install dependencies
```bash```
pip install xlwings openpyxl

### 2. xlwings addin install

### 3. Open Excel file
- Open `Calculator.xlsx` in Excel

### 4. Ensure xlwings Add-in is active
- Check Excel Add-ins tab

### 5. Run the application
- Click the **Compute** button
---
## How to Use

| Cell | Purpose |
|------|--------|
| B2 | Enter value for x |
| B3 | Enter value for y |
| B4 | Enter operation |
| B6 | Result (written by Python) |
| B7 | Status/Error message |


## Supported Operations

| Operation | Alias | Description |
|----------|------|-------------|
| SUM | ADD | x + y |
| PRODUCT | MULTIPLY | x × y |
| SUBTRACT | MINUS | x − y |
| DIVIDE | DIVISION | x ÷ y |
| POWER | POW | x ^ y |
| MODULO | MOD | x mod y |


## Error Handling

| Scenario | Status Message |
|----------|---------------|
| Empty inputs | ERROR: One or more inputs are empty |
| Non-numeric x or y | ERROR: x and y must be numeric values |
| Division by zero | ERROR: Division by zero is not allowed |
| Unsupported operation | ERROR: Unsupported operation |
| Unexpected error | UNEXPECTED ERROR |

## How It Works

1. User inputs values in Excel
2. Clicks "Compute"
3. Excel triggers Python via xlwings
4. Python validates inputs and performs calculation
5. Result is written back to Excel
