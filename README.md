# PDF Invoice to Excel Workflow Automation Demo

Demo: read invoices from PDF, check them against rules you set, and write Excel and Tally-ready import files.

[![Automate Invoices & Reports in Minutes! | Gentle Reminder Workflow Automation Demo](https://img.youtube.com/vi/DJBLUuWi6Ag/maxresdefault.jpg)](https://youtu.be/DJBLUuWi6Ag "Watch the demo on YouTube")

*Watch the walkthrough video: [Automate Invoices & Reports in Minutes! | Gentle Reminder Workflow Automation Demo](https://youtu.be/DJBLUuWi6Ag)*

---

## What It Does

- **Extracts structured fields from fixed-format PDFs**: Reads header details (invoice number, dates, payment terms), party details (buyer and vendor names, GSTINs, state codes), line items (descriptions, HSN/SAC codes, quantities, rates, amounts), and tax breakdowns.
- **Runs verification checks you define**: Validates calculations, tax arithmetic, and tax identifier structure against rule sets before writing results.
- **Flags discrepancies by name**: Points directly to specific failing checks (e.g. line math flag, tax mismatch flag) so discrepancies can be inspected immediately.
- **Writes consolidated Excel registers**: Produces structured workbooks including an Executive Dashboard with KPI cards, cross-vendor registers, and an item-level procurement price catalog.
- **Writes accounting import files**: Generates a Tally-ready purchase-voucher import file (XML), along with standard import files for Zoho Books and QuickBooks Online.

---

## The Set Checks

This demo evaluates each invoice against four verification checks:

1. **Line Item Math**: Verifies that `Quantity × Rate = Line Amount` for each line item within allowable rounding tolerance.
2. **Tax Math**: Verifies that `SubTotal + Total Tax = Grand Total`.
3. **GSTIN Format**: Checks that vendor and buyer tax numbers match the standard 15-character alphanumeric format.
4. **State Code Match**: Checks that the first two digits of the vendor GSTIN match the vendor's declared state code.

> **Honest note on verification:** It runs only the checks you set. Anything outside them is not checked.

---

## Example Result

In this demo dataset:
- **50 of 50 sample invoices passed all set checks.**

When an invoice contains an error, the audit report records the exact failure by name rather than failing silently. For example:
- A calculation variance is marked `Line Item Math: FLAGGED`.
- A tax difference is marked `SubTotal + Tax Match: FLAGGED`.
- An inconsistent state prefix is marked `State Code Match: MISMATCH`.

> **Sample data notice:** All invoices, company names, addresses, bank details, and tax identifiers in this repository are synthetic sample data generated specifically for demonstration purposes.

---

## Quick Start

### 1. Prerequisites
- Python 3.8 or higher

### 2. Clone the repository
```bash
git clone https://github.com/automategentlereminder/pdf-excel-workflow-automation-demo.git
cd pdf-excel-workflow-automation-demo
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```

### 4. Run the demo
```bash
python run_demo.py
```

### Where outputs land
The script processes the 50 PDF invoices in `sample-invoices/` and writes all output files directly to `sample-output/`:
- `sample-output/output.csv` — Master flat export of all extracted tags
- `sample-output/Management_Reports/Audit_Compliance_Report.xlsx` — Detailed check-by-check verification report
- `sample-output/Management_Reports/Executive_Invoice_Dashboard.xlsx` — Multi-tab register and KPI summary
- `sample-output/Management_Reports/Procurement_Price_Catalog.xlsx` — Item-level price variance catalog
- `sample-output/Accounting_Imports/Tally_Purchase_Vouchers.xml` — Tally-ready import file for Day Book
- `sample-output/Accounting_Imports/Zoho_Bills_Import.csv` — Zoho Books bill import file
- `sample-output/Accounting_Imports/QuickBooks_Bills_Import.csv` — QuickBooks Online batch bill import file
- `sample-output/By_Vendor/` — Subfolders with dedicated registers per vendor

---

## Configuration

Each script and template includes a top-level `CONFIG` block. Static adjustments can be made directly in this block without modifying parsing or extraction logic:

```python
# Example: top of run_demo.py
CONFIG = {
    "INPUT_DIR": "sample-invoices",       # Where PDFs are located
    "OUTPUT_DIR": "sample-output",        # Where outputs land
    "DETECTION_THRESHOLD": 70,            # Match score required
}
```

```python
# Example: top of output.py
CONFIG = {
    "DEFAULT_OUTPUT_DIR": "sample-output",
    "LINE_MATH_TOLERANCE": 1.0,           # Max allowable line math rounding difference
    "TAX_MATH_TOLERANCE": 2.0,            # Max allowable tax rounding difference
}
```

Vendor templates in `templates/<VENDOR>/` also provide top-level `CONFIG` dictionaries with vendor identifiers, keywords, and default terms.

---

## Repository Layout

```
pdf-excel-workflow-automation-demo/
├── .gitignore
├── LICENSE
├── README.md
├── requirements.txt
├── run_demo.py                     # Command-line runner orchestrating the demo
├── output.py                       # Generates Excel workbooks and accounting import files
├── tags.txt                        # Master list of 26 extraction fields
├── templates/                      # Rule-based vendor detection and extraction templates
│   ├── ACME/
│   │   ├── detect.py
│   │   └── extract.py
│   ├── GLOBEX/
│   │   ├── detect.py
│   │   └── extract.py
│   ├── INITECH/
│   │   ├── detect.py
│   │   └── extract.py
│   ├── STARK/
│   │   ├── detect.py
│   │   └── extract.py
│   └── WAYNE/
│       ├── detect.py
│       └── extract.py
├── sample-invoices/                # 50 synthetic sample invoice PDFs across 5 vendors
│   ├── ACM-25-26-0013.pdf
│   └── ...
└── sample-output/                  # Pre-generated sample outputs ready to inspect
    ├── output.csv
    ├── Accounting_Imports/
    │   ├── QuickBooks_Bills_Import.csv
    │   ├── Tally_Purchase_Vouchers.xml
    │   └── Zoho_Bills_Import.csv
    ├── By_Vendor/
    └── Management_Reports/
        ├── Audit_Compliance_Report.xlsx
        ├── Executive_Invoice_Dashboard.xlsx
        └── Procurement_Price_Catalog.xlsx
```

---

## Limits

- **Fixed formats**: Built for regular, fixed invoice layouts that a business receives from recurring vendors. When a vendor introduces a new layout or format, a corresponding template rule is created or adjusted.
- **Execution environment**: Runs on your own computer by default and connects only where a workflow needs it.

---

## About Gentle Reminder

Built by Gentle Reminder, custom PDF and Excel workflow automation for businesses.

- Website: [gentlereminder.in](https://gentlereminder.in/)
- Workflow Automation: [gentlereminder.in/workflow-automation/](https://gentlereminder.in/workflow-automation/)
- Further reading:
  - [What set checks mean in invoice automation](https://gentlereminder.in/blog/what-set-checks-mean-in-invoice-automation/)
  - [Invoice automation: speed vs verification](https://gentlereminder.in/blog/invoice-automation-speed-vs-verification/)

---

## Contact

For questions or custom workflow automation requests:
- Email: [hello@gentlereminder.in](mailto:hello@gentlereminder.in)
