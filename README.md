# PDF Invoice Modifier

This tool automates the modification of PDF invoices by reducing "Unit Price" by 40%, recalculating totals, and removing transport costs.

## Prerequisites
1.  **Python 3.x** installed.
2.  **PyMuPDF** library installed:
    ```bash
    pip install pymupdf
    ```

## How to Use

1.  **Prepare Input**:
    -   Place your original PDF invoice file in the `input/` folder.
    -   Rename it to `invoice.pdf` (or update `INPUT_PATH` in `main.py`).

2.  **Run the Script**:
    Open a terminal in this folder and run:
    ```bash
    py main.py
    ```

3.  **Check Output**:
    -   The modified PDF will be saved in the `output/` folder as `invoice_modified.pdf`.

## Configuration
If your invoice uses different headers or languages, open `main.py` and edit the following variables at the top:

```python
HEADER_UNIT_PRICE = "Unit Price"
HEADER_TOTAL = "Total"
KEYWORD_TRANSPORT = "Transport"
KEYWORD_INVOICE_TOTAL = "Invoice Total"
```

## Troubleshooting
-   **Columns not detected**: Ensure the headers in your PDF match the keywords above.
-   **Values not changing**: The script uses a regex for European currency format (`1.234,56`). If your invoice uses `1,234.56`, update `CURRENCY_REGEX` and the parsing logic in `main.py`.
