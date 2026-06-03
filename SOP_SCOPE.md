# SOP Scope Alignment

This project automates the pre-accounting preparation described in `SOP-P2P-CO-PRO-003 Arrendamientos (Rentas) CO Rev3`.

## Automated scope

The app covers the logic before manual posting in Lucy/SAP:

- Step 1: preparation and analysis of information
- Step 3: preparation of accounting data for Lucy
- Step 5: prorated VAT calculation
- Exception handling for vendor distribution cases like `33%` and `50%`

## What the app does

- Loads the source Excel files without manual editing
- Cleans and standardizes source fields
- Maps store, vendor, and CECO
- Checks contract status by CECO
- Determines rent type `RF` or `RV`
- Determines the accounting account
- Calculates `VAT_VW` and `VAT_VQ`
- Splits invoices by distribution rules when required
- Generates standardized accounting text
- Produces a single consolidated output file for downstream operation

## What the app does not do

- Post entries in Lucy
- Post entries in SAP
- Validate tax withholdings in SAP
- Execute compensations such as `F-44`

These actions remain good candidates for later RPA automation.

## Output objective

The output file is the single downstream input for operational use:

- One row per accounting line
- Includes vendor, CECO, account, amount, VAT split, and text
- No additional calculations should be required by the future bot

Target output:

- `output_lucy_ready.xlsx`
