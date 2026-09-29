# P2P Invoice Matching and Aging Tracker

A practice project built in Excel with sample data to show Accounts Payable (Procure to Pay) work.

## What it does
- 3-way match of PO, GRN and invoice
- Flags exceptions: quantity mismatch, price mismatch, duplicate invoice, missing PO, no GRN, PO not found, vendor mismatch
- Invoice aging report (1-30, 31-60, 61-90 days) using a pivot table
- Vendor statement reconciliation

## Sheets
Summary, PO_Data, GRN_Data, Invoice_Data, Matching, Aging_Report, Vendor_Recon

## Excel skills used
VLOOKUP, IFERROR, nested IF, COUNTIF (expanding range), date subtraction, pivot table

Note: All data is made-up sample data.
