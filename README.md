# RFQ to quote draft

A small, deterministic proof that turns a PDF request for quotation into an Excel quote draft without guessing prices. Every RFQ line and catalogue entry is synthetic. This is independent work, not client delivery, and it is not an n8n production workflow.

Live page: https://sandersaxbuilds.github.io/rfq-quote-synthetic-proof/

## What the script does

1. Renders the synthetic text request `source-rfq.txt` to `sample-rfq.pdf`.
2. Parses the generated PDF bytes again and writes the extracted text to `extracted-rfq.txt`.
3. Matches each RFQ line to `catalog.csv` by exact SKU only.
4. Gives matched SKUs the approved catalogue price.
5. Marks an unknown SKU `REVIEW_REQUIRED`, with no price and a note.
6. Writes `result.json` and a real Excel workbook, `quote-draft.xlsx`.

With the included data, three of four lines match and are priced at USD 467.30 in total. `SKU-9999` is left unpriced for review.

## Rules

- Unknown products are never swapped for a similar item.
- No price is invented for an unmatched line.
- The workbook is a draft for review, not an accepted quotation.

## Run it

Windows with Microsoft Excel installed:

```powershell
powershell -ExecutionPolicy Bypass -File .\build-proof.ps1
```

The script reads `source-rfq.txt` and `catalog.csv` from the same folder and rewrites `sample-rfq.pdf`, `extracted-rfq.txt`, `result.json` and `quote-draft.xlsx`. The committed outputs come from a previous run of the same script.

## Files

| File | Role |
|---|---|
| `source-rfq.txt` | Synthetic RFQ text, input |
| `catalog.csv` | Approved catalogue with SKU, description, unit price and currency, input |
| `build-proof.ps1` | The whole pipeline |
| `sample-rfq.pdf` | Rendered RFQ, output |
| `extracted-rfq.txt` | Text read back from the PDF, output |
| `result.json` | Line-by-line match result, output |
| `quote-draft.xlsx` | Excel quote draft, output |
| `index.html` | GitHub Pages summary |
