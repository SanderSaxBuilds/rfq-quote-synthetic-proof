# RFQ to quote synthetic proof

This is new independent proof prepared for a possible RFQ-to-quote paid test. It is not client delivery and it does not claim prior paid n8n work.

## What it proves

1. A synthetic text RFQ is rendered to `sample-rfq.pdf`.
2. The generated PDF bytes are parsed again and the text objects are extracted into `extracted-rfq.txt`.
3. RFQ lines are matched only by exact SKU against `catalog.csv`.
4. Approved SKUs receive the approved catalogue price.
5. An unknown SKU is marked `REVIEW_REQUIRED` and receives no guessed price.
6. The result is written to `result.json` and a real Excel workbook at `quote-draft.xlsx`.

## Safety and scope

- All RFQ and catalogue data is synthetic.
- Unknown products are never silently substituted.
- No price is invented for an unmatched item.
- The workbook is a draft for review, not an accepted quotation.
- The proof uses a deterministic local script. It does not claim that an n8n production workflow has been delivered.

## Run locally

On Windows with Microsoft Excel installed:

```powershell
powershell -ExecutionPolicy Bypass -File .\build-proof.ps1
```

The generated artifacts are `sample-rfq.pdf`, `extracted-rfq.txt`, `result.json`, and `quote-draft.xlsx`.
