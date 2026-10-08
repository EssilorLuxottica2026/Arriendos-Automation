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
- Office invoices: when `Tienda` contains `OFICINA`, office-specific accounts are read from the `*_OFICINA` entries in `Diccionario_Conceptos_Simple.xlsx`; the exported concepts keep the common siglas
- VAT proration into `VAT_VW` and `VAT_VQ`
- Distribution handling for special vendors
- Consolidated output generation for downstream operation

VAT proration uses the invoice vendor together with the resolved store first,
then vendor + resolved CECO, store, CECO, a generic vendor rule, and finally the
macro database. Generic vendor rules must have no store or CECO: a percentage
assigned to another location must not override the current location. VW and VQ
are taken from the same rule (missing VQ is calculated as `1 - VW`), and the
selected source is recorded in the execution log.
For partial proration, the posting lines retain the nondeductible VAT remainder
(`VAT_VQ`) with posting key `50` and `VQ`. They also use the rounded deductible
VAT amount twice: posting key `40` with `VW`, and an additional posting key `50`
with `VQ` for the same amount (`vat_vq_reversal`). The additional reversal does
not replace the nondeductible VAT line.
Full VW or full VQ proration keeps the existing behavior without separate VAT
adjustment lines.

See `SOP_SCOPE.md` for the project scope summary.

Concept review selections contain `concept|account`. Saving learned phrases
separates these values, verifies the account against the existing dictionary
row, and updates that row rather than creating a combined-name concept.
Unknown concepts or mismatched accounts require reloading the review; they do
not create incomplete dictionary rules.

Dictionary rows have independent `Concepto`, `GL Account`, `Frases/palabras`,
and `Sigla` values. Different concepts may share an account while keeping
different output abbreviations (for example AIRE, ASEO, and VIGILANCIA).
Concept names must be unique. RENTA uses `SEGUN_CONTRATO`; GC, GC V, and
GC VARIABLE keep the SELL MEDIA-based account rule. Office concepts use their
explicit `*_OFICINA` dictionary accounts while retaining the output abbreviation
of their common counterpart. Avoid assigning the same phrase to different
concepts: the longest matching phrase wins, with dictionary row order breaking
ties.

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
