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

## Invoice control for the bot

Each successful processing run saves `control_facturas_bot.json` in its own
output directory. Run directories include seconds and microseconds so a new run
does not overwrite an earlier control file or the bot's progress.
The results page and command-line runner expose the JSON and a bot ZIP package.
The bot package contains only the JSON and the linked Allocated Costs CSVs,
keeping their relative directory structure. It excludes invoice XML supports,
header CSVs and manual-style invoice CSVs. Extract the complete package before
using it; every `csv.allocated_costs` path resolves from the extracted JSON.
The other file references remain in the JSON for preparation traceability and
refer to the full output directory, not files included in the bot ZIP.
XML validation and local support copies are preserved, as are the separate
invoice CSV downloads. The bot does not accept PDF supports.

The versioned JSON contains a `facturas` list with one entry per processed folio:

- `folio`, `vendor`, `vendors`, and total `importe`.
- `ruta`, `estado`, `lista_para_bot`, `documento_sap`, UTC processing/update
  dates, `error`, and `pasos_confirmados`.
- `partes`: each posting's vendor, store (`tienda`), CeCo, profit center, base,
  VAT, amount, currency, invoice date, and document type.
- Each part links its copied support (`archivo_factura`), original source path,
  header CSV, Allocated Costs CSV, and manual-style invoice CSV.
  `allocated_costs` also contains the exact exported rows, without the internal
  posting index. CSV and copied-support paths are relative to the JSON directory.
- Each part reserves `identificador_lucy`, `barcode`, `documento_sap`,
  `copia_lucy`, and `compensacion_sap`. Unknown values remain JSON `null`;
  preparation never confirms a Lucy copy, SAP posting, or compensation.

Only `single_vendor` is enabled. A folio with several postings for the same vendor
keeps separate parts and CSVs. Folios distributed across multiple vendors retain
their existing accounting exports but have `ruta: null`,
`estado: "pendiente_ruta"`, and `lista_para_bot: false`. Missing invoice supports
are reported as `pendiente_archivo`, also preventing bot execution.
`preparada` confirms only the preparation step, not execution by the bot.

### Preparation validation (schema version 2)

Before saving the control, the pipeline verifies:

- A nonblank folio, vendor and store; a valid CeCo using the existing rules;
  profit center, invoice date, COP currency and a positive total equal to
  base plus VAT. Base and VAT must not be negative.
- The support is a readable invoice/credit-note/debit-note XML. Its folio,
  document type, invoice date and available `ParentDocumentID` must agree
  with the prepared record. A folio linked to multiple XMLs requires review.
- For an XML shared across parts, the sum of gross amounts and VAT must match
  the XML's extracted amounts, using the existing whole-peso rounding rules.
  Gross amounts are checked before applied discounts, not against the net
  payable amount, so discount preparation is preserved.
- The header, Allocated Costs and manual-style CSVs exist inside the output
  directory and are linked to a single part. The header and Allocated Costs
  CSV contents must match the prepared rows. The manual CSV must retain the
  correct folio, total and allocated costs section. Empty allocations,
  incorrect assignments, missing accounts and mismatched profit centers
  require review.

An uncertain or incorrect relationship is saved with `pendiente_revision`,
an explicit `error`, and `lista_para_bot: false`; it also appears in the results
warnings. Missing required files prevent publication of a successful control.
Successful checks append `validacion_preparacion_completada` to confirmed steps.
The JSON includes per-invoice `validaciones` and a `resumen_validacion` with
validated and pending counts. Validation is scoped to preparation: it is not
a substitute for checks immediately before the future bot executes.

Duplicate handling is deliberately deferred. `verificacion_duplicados` is
`no_ejecutada`, and `lista_para_bot` means only that the data, XML and CSV
preparation passed. Neither prior JSON controls nor an external posting list
are checked. `FACTURAS_CONTABILIZADAS` remains exclusively the account catalog
and must not be interpreted as evidence that an invoice was posted.

Optional invoice input columns `Lucy_ID` / `Identificador_Lucy` / `ID_Lucy` /
`Document_ID` and `Barcode` / `Bar_Code` / `Codigo_de_barras` / `Codigo_barras`
are carried into the corresponding parts. Text identifiers retain leading zeros
(Excel cells must store them as text). If these columns are absent, the bot can
fill the identifiers later.
For XML inputs and matched XML supports, `barcode` is read from the outer
document's `cbc:ParentDocumentID`, including `AttachedDocument` containers,
before extracting the embedded invoice or credit/debit note. Its text value is
preserved (for example `004092`); this XML value takes precedence over the table
column. If the XML has no `ParentDocumentID`, the optional table value is retained
or the JSON contains `null`. Neither the container's `cbc:ID` nor the invoice
folio is used as a substitute barcode.

Invoices excluded by current validation/review rules are not in the executable
`facturas` list. Available omission reasons are stored separately in
`incidencias_procesamiento`; the existing reports and warnings remain available.
The JSON is a preparation snapshot, not a global duplicate-posting registry:
the future bot must check Lucy/SAP before executing invoices from another run.

## Out of scope

- Posting in Lucy
- Posting in SAP
- SAP withholding validation
- Compensations like `F-44`
