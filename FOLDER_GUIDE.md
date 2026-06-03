# Folder Guide

This repo keeps the runtime code easy to find while separating operational data, support material, and maintenance scripts.

## Code

- `app.py`: Flask entrypoint
- `pipeline/`: processing logic and PDF/XML parsing
- `templates/`: HTML views
- `static/`: CSS and browser assets
- `scripts/`: maintenance and workbook preparation scripts

## Working Data

- `data/uploads/`: files submitted to the app during a batch
- `data/raw/pdfs/`: invoice PDFs used during analysis
- `data/raw/xmls/`: electronic invoice XML supports
- `data/normalized_inputs/`: normalized workbooks created by helper scripts
- `data/standardized_workbooks/`: standardized macro workbook artifacts
- `data/macro_export/`: macro-oriented exports when generated
- `data/output/`: all generated outputs from the pipeline

## Support Material

- `docs/reference/`: notes and reference text kept for onboarding
- `docs/business/`: business validation workbook and guide
- `support/media/`: media and other non-runtime reference material

## Scratch / Temporary

- `scratch/`: temporary files and experiments that should not be treated as source data

## Cleanup Rules

- Keep generated outputs and scratch files out of Git.
- Keep code changes isolated from working files used by the current process.
- When adding a new artifact, decide first whether it is source, support, or temporary.
