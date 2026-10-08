# Lease Accounting Web App

Flask web app to upload lease accounting source files, process invoices, and export a Lucy/SAP-ready output.

This project follows the pre-accounting scope from the SOP `Arrendamientos - Facturas Manuales` and is designed to centralize the Excel-based preparation work before manual or automated entry into Lucy/SAP.

## Run locally

```bash
python -m pip install -r requirements.txt
python app.py
```

Then open `http://127.0.0.1:5000`.

## Required uploads

- Invoices file: `.xlsx`, `.xls`, or `.csv`
- `CONTROL ARRI ADMON.xlsx`
- `Contratos con condiciones.xlsx`
- `PRORATEO.xlsx`
- `Cuadro de distribucion 2023.xls`

Optional:

- `FACTURAS CONTABILIZADAS ARRIENDOS.xlsx`

## Invoice columns

Expected names are flexible, but the app should be able to detect these concepts:

- `Invoice_ID`
- `Vendor`
- `Amount`
- `VAT`
- `Store`
- `Reference`
- `Invoice_Date`
- `CECO`

Use `sample_invoices.csv` as a template if needed.

## Scope covered

- Preparation and analysis of input information
- Cross-check of store, vendor, CECO, and contracts
- Determination of `RF` and `RV`
- Accounting account assignment
- VAT proration into `VAT_VW` and `VAT_VQ`
- Distribution handling for special vendors
- Consolidated output generation for downstream operation

See `SOP_SCOPE.md` for the project scope summary.

## Project structure

- `app.py`: thin wrapper that keeps the old entrypoint working
- `src/lease_accounting/`: main application package
- `src/lease_accounting/web.py`: Flask routes and upload flow
- `src/lease_accounting/pipeline/`: pipeline and support parsing logic
- `pipeline/`: compatibility layer for older imports

## Out of scope

- Posting in Lucy
- Posting in SAP
- SAP withholding validation
- Compensations like `F-44`
