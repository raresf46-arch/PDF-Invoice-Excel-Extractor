# PDF Invoice → Excel Extractor (with automatic checks)

Reads a folder of invoice PDFs from **different suppliers, in different layouts and languages (EN / RO)**,
and produces one Excel workbook. It also **checks the math on every invoice**, so errors are caught before payment.

![Before and after](docs/before_after.png)

## What it does

1. Detects which supplier layout each PDF uses
2. Extracts the invoice number, date, supplier, subtotal, VAT, total and currency
3. Extracts every line item from the invoice table
4. Handles both `1,234.56` and `1.234,56` number formats
5. Validates each invoice:
   - sum of line items = subtotal?
   - subtotal + VAT = total?
6. Highlights problem invoices in red

## Output (one Excel file, 3 sheets)

| Sheet | Contents |
|---|---|
| **Invoices** | One row per invoice: number, date, supplier, subtotal, VAT, total, currency |
| **Line Items** | Every product line from every invoice |
| **Checks** | Validation per invoice: green **OK** or red **CHECK: total off by +50.00** |

## Run it

```bash
pip install pdfplumber pandas openpyxl reportlab
python make_sample_invoices.py                 # creates 9 demo invoices (2 layouts, 1 with an error)
python extract_invoices.py sample_invoices -o output/invoices.xlsx
```

```
  factura_DMC780.pdf           Romanian factura   3 line items
  ...
  invoice_NTS-2405.pdf         English invoice    3 line items

9 invoices, 27 line items extracted. 1 invoice(s) need checking.
```

## Adding a new supplier

Each layout is one `Layout(...)` profile at the top of `extract_invoices.py`, containing the patterns for each field
and the table column names. A new supplier format is about 10 lines of configuration; the code itself doesn't change.

## Tech

Python · pdfplumber · regex · pandas · openpyxl
