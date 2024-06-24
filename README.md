# Invoice Management System

## Overview
This Python script implements an Invoice Management System using tkinter for the GUI, pandas for data handling, openpyxl for Excel operations, and pdfkit for PDF generation.

## Features
- **Create Invoices**: Enter customer details, invoice date, amount, hourly rate, hours booked, and description to generate invoices.
- **View and Manage Invoices**: Lists all invoices with details like invoice number, customer name, amount, due date, and status. Includes functionalities to mark invoices as paid and search for specific invoices.
- **Generate PDF Invoices**: Capable of generating PDF invoices from Excel templates, although this feature is currently commented out in the provided script.

## Dependencies
- **Python Libraries**:
  - `os`
  - `datetime`
  - `timedelta` from `datetime`
  - `openpyxl`
  - `pandas`
  - `tkinter` (for GUI)
  - `ttk` from `tkinter` (for themed widgets)
  - `messagebox` from `tkinter`
  - `Canvas` from `tkinter`
  - `load_workbook` from `openpyxl`
  - `pdfkit`

## Usage
1. Clone the repository and navigate to the project directory.
2. Run the application with `python invoice_app.py`.
