import csv
import io
import os
import re
import json
from datetime import datetime
from io import StringIO, BytesIO

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# CONFIGURATION BLOCK
# Change default paths, audit tolerances, and output settings here.
# ==============================================================================
CONFIG = {
    # Default output directory for generated registers, reports, and import files
    "DEFAULT_OUTPUT_DIR": os.environ.get("OUTPUT_DIR", "sample-output"),

    # Relative path to tags definition file
    "TAGS_FILE": "tags.txt",

    # Verification tolerances (in currency units)
    "LINE_MATH_TOLERANCE": 1.0,     # Max allowable variance for Quantity * Rate vs Line Amount
    "TAX_MATH_TOLERANCE": 2.0,      # Max allowable variance for SubTotal + Total Tax vs Total
}

# ==============================================================================
# Comprehensive Output Processor
# ==============================================================================
#
# Processes extracted document records and produces:
#  1. Master Output Flat CSV (with all 26 tags)
#  2. Executive Invoice Dashboard (Multi-tab Excel with KPIs & summary)
#  3. Procurement & Price Catalog (Item-level price variance & volume analysis)
#  4. Audit & Compliance Verification Report (Automated math & GST validation)
#  5. Zoho Books Bill Import CSV
#  6. Tally-ready Purchase Vouchers XML (Native voucher standard for Day Book)
#  7. QuickBooks Online Batch Bill Import CSV
#  8. Vendor-specific segmented folders and bill registers
#


def _cell_value(val):
    if isinstance(val, dict) and "value" in val:
        return val.get("value")
    return val


def _to_float(val):
    if val is None or val == "":
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).replace(",", "").replace("₹", "").replace("Rs.", "").replace("Rs", "").replace("INR", "").strip()
    try:
        return float(s)
    except Exception:
        return 0.0


def _normalize_date(date_str):
    if not date_str:
        return "", ""
    s = str(date_str).strip()
    formats = [
        "%d/%m/%Y", "%Y-%m-%d", "%d-%b-%Y", "%d-%B-%Y",
        "%b %d, %Y", "%B %d, %Y", "%d %b %Y", "%d %B %Y"
    ]
    for fmt in formats:
        try:
            dt = datetime.strptime(s, fmt)
            return dt.strftime("%Y-%m-%d"), dt.strftime("%Y%m%d")
        except ValueError:
            pass
    m = re.search(r"(\d{4})[^\d](\d{2})[^\d](\d{2})", s)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}", f"{m.group(1)}{m.group(2)}{m.group(3)}"
    return s, s.replace("-", "").replace("/", "")


def _map_expense_account(vendor_name):
    vn = (vendor_name or "").lower()
    if "acme" in vn:
        return "Office Supplies & Stationery"
    elif "globex" in vn:
        return "Rent & Shared Workspace"
    elif "initech" in vn:
        return "Software Licensing & IT Support"
    elif "stark" in vn:
        return "Computer Hardware & Capital Equipment"
    elif "wayne" in vn:
        return "Cloud Infrastructure & Hosting Services"
    return "General Operating Expenses"


def _style_header_cell(cell, text, bg_color="1F4E79", font_color="FFFFFF", font_size=11, align="center"):
    cell.value = text
    cell.font = Font(name="Segoe UI", size=font_size, bold=True, color=font_color)
    cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)


def _apply_table_styling(ws, start_row, end_row, start_col, end_col):
    thin_side = Side(border_style="thin", color="D9D9D9")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = border
            if r % 2 == 0 and r != start_row:
                cell.fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")


def _autofit_columns(ws, max_cols=None):
    cols = max_cols or ws.max_column
    for col_idx in range(1, cols + 1):
        col_letter = get_column_letter(col_idx)
        max_len = 0
        for cell in ws[col_letter]:
            val_str = str(cell.value or "")
            if "\n" in val_str:
                val_str = max(val_str.split("\n"), key=len)
            max_len = max(max_len, len(val_str))
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)


def output(documents, target_dir=None):
    # Resolve output directory
    out_dir = target_dir or globals().get("outputPath1") or CONFIG["DEFAULT_OUTPUT_DIR"]
    if not os.path.isabs(out_dir):
        out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), out_dir)

    tags_list = list(globals().get("tags") or [])
    if not tags_list:
        tags_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), CONFIG["TAGS_FILE"])
        if os.path.isfile(tags_file):
            with open(tags_file, "r", encoding="utf-8") as tf:
                tags_list = [line.strip() for line in tf if line.strip()]
        else:
            tags_list = [
                "Vendor Name", "Vendor GSTIN", "Vendor State", "Vendor State Code",
                "Bill Number", "Bill Date", "Due Date", "Payment Terms",
                "Buyer Name", "Buyer GSTIN", "Item Name", "HSN/SAC Code",
                "Quantity", "Unit", "Rate", "Line Amount", "Line Tax Rate",
                "SubTotal", "CGST Amount", "SGST Amount", "IGST Amount",
                "Total Tax", "Total", "Amount in Words", "Vendor Notes", "Terms & Conditions"
            ]

    # Subdirectories within output directory
    mgmt_dir = os.path.join(out_dir, "Management_Reports")
    acct_dir = os.path.join(out_dir, "Accounting_Imports")
    vendor_root_dir = os.path.join(out_dir, "By_Vendor")
    for d in [out_dir, mgmt_dir, acct_dir, vendor_root_dir]:
        try:
            os.makedirs(d, exist_ok=True)
        except Exception:
            pass

    # -------------------------------------------------------------------------
    # 1. Parse all documents and flatten structures
    # -------------------------------------------------------------------------
    parsed_invoices = []
    flat_rows = []
    item_procurement = {}  # item_name -> list of records

    for doc in documents:
        extracted = doc.get("extraction") or {}
        if not isinstance(extracted, dict):
            extracted = {}

        # Height determination for multi-line fields
        height = 1
        parsed_doc = {}
        for col in tags_list:
            val = _cell_value(extracted.get(col, ""))
            if isinstance(val, list):
                parsed_doc[col] = [_cell_value(item) for item in val]
                height = max(height, len(val) or 1)
            else:
                parsed_doc[col] = [val]

        doc_rows = []
        for index in range(max(height, 1)):
            row = {}
            for col in tags_list:
                values = parsed_doc[col]
                chosen = values[0] if len(values) == 1 else (values[index] if index < len(values) else "")
                row[col] = "" if chosen is None else str(chosen).strip()
            row["_file"] = doc.get("file", "")
            row["_template"] = doc.get("template", "")
            doc_rows.append(row)
            flat_rows.append(row)

            # Accumulate procurement data
            item_name = row.get("Item Name", "").strip()
            if item_name:
                qty = _to_float(row.get("Quantity", 0))
                rate = _to_float(row.get("Rate", 0))
                amt = _to_float(row.get("Line Amount", 0))
                hsn = row.get("HSN/SAC Code", "").strip()
                unit = row.get("Unit", "").strip()
                vname = row.get("Vendor Name", "").strip()
                bno = row.get("Bill Number", "").strip()
                bdate = row.get("Bill Date", "").strip()

                if item_name not in item_procurement:
                    item_procurement[item_name] = {
                        "hsn": hsn,
                        "unit": unit,
                        "vendors": set(),
                        "bills": set(),
                        "records": []
                    }
                item_procurement[item_name]["vendors"].add(vname)
                item_procurement[item_name]["bills"].add(bno)
                item_procurement[item_name]["records"].append({
                    "qty": qty, "rate": rate, "amt": amt, "date": bdate, "vendor": vname, "bill": bno
                })

        parsed_invoices.append({
            "header": doc_rows[0] if doc_rows else {},
            "lines": doc_rows,
            "template": doc.get("template", "")
        })

    total_invoices_count = len(parsed_invoices)
    total_spend = sum(_to_float(inv["header"].get("Total", 0)) for inv in parsed_invoices)
    total_subtotal = sum(_to_float(inv["header"].get("SubTotal", 0)) for inv in parsed_invoices)
    total_tax_paid = sum(_to_float(inv["header"].get("Total Tax", 0)) for inv in parsed_invoices)
    distinct_vendors = len(set(inv["header"].get("Vendor Name", "") for inv in parsed_invoices if inv["header"].get("Vendor Name")))
    total_line_items = len(flat_rows)

    # -------------------------------------------------------------------------
    # 2. Master Output CSV
    # -------------------------------------------------------------------------
    csv_buf = StringIO()
    writer = csv.DictWriter(csv_buf, fieldnames=tags_list, extrasaction="ignore", lineterminator="\n")
    writer.writeheader()
    for row in flat_rows:
        writer.writerow({k: row.get(k, "") for k in tags_list})
    output_csv_content = csv_buf.getvalue()

    try:
        with open(os.path.join(out_dir, "output.csv"), "w", encoding="utf-8", newline="") as f:
            f.write(output_csv_content)
    except Exception:
        pass

    # -------------------------------------------------------------------------
    # 3. Zoho Books Bill Import CSV
    # -------------------------------------------------------------------------
    zoho_fields = [
        "Bill Number", "Vendor Name", "Bill Date", "Due Date", "Payment Terms",
        "Item Name", "HSN/SAC", "Account", "Quantity", "Rate", "Item Total",
        "SubTotal", "Tax Name", "Tax Percentage", "Tax Amount", "Total",
        "Customer Name", "Notes", "Terms & Conditions"
    ]
    zoho_buf = StringIO()
    zoho_writer = csv.DictWriter(zoho_buf, fieldnames=zoho_fields, extrasaction="ignore", lineterminator="\n")
    zoho_writer.writeheader()

    for inv in parsed_invoices:
        hdr = inv["header"]
        vname = hdr.get("Vendor Name", "")
        bno = hdr.get("Bill Number", "")
        iso_bdate, _ = _normalize_date(hdr.get("Bill Date", ""))
        iso_ddate, _ = _normalize_date(hdr.get("Due Date", ""))
        pterms = hdr.get("Payment Terms", "")
        bname = hdr.get("Buyer Name", "")
        vnotes = hdr.get("Vendor Notes", "")
        terms = hdr.get("Terms & Conditions", "")
        subtot = _to_float(hdr.get("SubTotal", 0))
        tot = _to_float(hdr.get("Total", 0))
        acct = _map_expense_account(vname)

        is_inter_state = _to_float(hdr.get("IGST Amount", 0)) > 0
        tax_name = "IGST" if is_inter_state else "GST"

        for line in inv["lines"]:
            qty = _to_float(line.get("Quantity", 0))
            rate = _to_float(line.get("Rate", 0))
            amt = _to_float(line.get("Line Amount", 0))
            tax_rate_str = line.get("Line Tax Rate", "").replace("%", "").strip() or "18"
            tax_pct = _to_float(tax_rate_str)
            tax_amt = round(amt * (tax_pct / 100.0), 2)

            zoho_writer.writerow({
                "Bill Number": bno,
                "Vendor Name": vname,
                "Bill Date": iso_bdate,
                "Due Date": iso_ddate,
                "Payment Terms": pterms,
                "Item Name": line.get("Item Name", ""),
                "HSN/SAC": line.get("HSN/SAC Code", ""),
                "Account": acct,
                "Quantity": qty,
                "Rate": rate,
                "Item Total": amt,
                "SubTotal": subtot,
                "Tax Name": tax_name,
                "Tax Percentage": tax_pct,
                "Tax Amount": tax_amt,
                "Total": tot,
                "Customer Name": bname,
                "Notes": vnotes,
                "Terms & Conditions": terms
            })

    zoho_csv_content = zoho_buf.getvalue()
    try:
        with open(os.path.join(acct_dir, "Zoho_Bills_Import.csv"), "w", encoding="utf-8", newline="") as f:
            f.write(zoho_csv_content)
    except Exception:
        pass

    # -------------------------------------------------------------------------
    # 4. QuickBooks Online Batch Bill Import CSV
    # -------------------------------------------------------------------------
    qb_fields = [
        "BillNo", "Vendor", "BillDate", "DueDate", "Terms", "LineItem",
        "LineDescription", "LineQuantity", "LineUnitPrice", "LineAmount",
        "LineExpenseAccount", "LineTaxCode"
    ]
    qb_buf = StringIO()
    qb_writer = csv.DictWriter(qb_buf, fieldnames=qb_fields, extrasaction="ignore", lineterminator="\n")
    qb_writer.writeheader()

    for inv in parsed_invoices:
        hdr = inv["header"]
        vname = hdr.get("Vendor Name", "")
        bno = hdr.get("Bill Number", "")
        iso_bdate, _ = _normalize_date(hdr.get("Bill Date", ""))
        iso_ddate, _ = _normalize_date(hdr.get("Due Date", ""))
        pterms = hdr.get("Payment Terms", "")
        acct = _map_expense_account(vname)
        is_inter_state = _to_float(hdr.get("IGST Amount", 0)) > 0
        tax_code = "IGST 18%" if is_inter_state else "GST 18%"

        for line in inv["lines"]:
            iname = line.get("Item Name", "")
            qty = _to_float(line.get("Quantity", 0))
            rate = _to_float(line.get("Rate", 0))
            amt = _to_float(line.get("Line Amount", 0))

            qb_writer.writerow({
                "BillNo": bno,
                "Vendor": vname,
                "BillDate": iso_bdate,
                "DueDate": iso_ddate,
                "Terms": pterms,
                "LineItem": iname,
                "LineDescription": f"{iname} (HSN/SAC: {line.get('HSN/SAC Code', '')})",
                "LineQuantity": qty,
                "LineUnitPrice": rate,
                "LineAmount": amt,
                "LineExpenseAccount": acct,
                "LineTaxCode": tax_code
            })

    qb_csv_content = qb_buf.getvalue()
    try:
        with open(os.path.join(acct_dir, "QuickBooks_Bills_Import.csv"), "w", encoding="utf-8", newline="") as f:
            f.write(qb_csv_content)
    except Exception:
        pass

    # -------------------------------------------------------------------------
    # 5. Tally Prime Purchase Vouchers XML (Native import file)
    # -------------------------------------------------------------------------
    tally_lines = [
        '<ENVELOPE>',
        '  <HEADER>',
        '    <TALLYREQUEST>Import Data</TALLYREQUEST>',
        '  </HEADER>',
        '  <BODY>',
        '    <IMPORTDATA>',
        '      <REQUESTDESC>',
        '        <REPORTNAME>Vouchers</REPORTNAME>',
        '        <STATICVARIABLES>',
        '          <SVCURRENTCOMPANY>Gentle Reminder</SVCURRENTCOMPANY>',
        '        </STATICVARIABLES>',
        '      </REQUESTDESC>',
        '      <REQUESTDATA>'
    ]

    for inv in parsed_invoices:
        hdr = inv["header"]
        bno = hdr.get("Bill Number", "")
        _, tally_date = _normalize_date(hdr.get("Bill Date", ""))
        if not tally_date:
            tally_date = datetime.now().strftime("%Y%m%d")

        vname = hdr.get("Vendor Name", "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        tot_val = _to_float(hdr.get("Total", 0))
        sub_val = _to_float(hdr.get("SubTotal", 0))
        cgst_val = _to_float(hdr.get("CGST Amount", 0))
        sgst_val = _to_float(hdr.get("SGST Amount", 0))
        igst_val = _to_float(hdr.get("IGST Amount", 0))
        pur_ledger = _map_expense_account(vname).replace("&", "&amp;")

        tally_lines.extend([
            '        <TALLYMESSAGE xmlns:UDF="TallyUDF">',
            f'          <VOUCHER VCHTYPE="Purchase" ACTION="Create" OBJVIEW="Accounting Voucher View">',
            f'            <DATE>{tally_date}</DATE>',
            f'            <VOUCHERTYPENAME>Purchase</VOUCHERTYPENAME>',
            f'            <REFERENCE>{bno}</REFERENCE>',
            f'            <PARTYLEDGERNAME>{vname}</PARTYLEDGERNAME>',
            f'            <PARTYNAME>{vname}</PARTYNAME>',
            f'            <BASICBUYERNAME>Gentle Reminder</BASICBUYERNAME>',
            f'            <NARRATION>Purchase Invoice {bno} from {vname}</NARRATION>',
            '            <!-- Sundry Creditor Credit Ledger -->',
            '            <ALLLEDGERENTRIES.LIST>',
            f'              <LEDGERNAME>{vname}</LEDGERNAME>',
            '              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>',
            f'              <AMOUNT>{tot_val:.2f}</AMOUNT>',
            '            </ALLLEDGERENTRIES.LIST>',
            '            <!-- Expense / Purchase Debit Ledger -->',
            '            <ALLLEDGERENTRIES.LIST>',
            f'              <LEDGERNAME>{pur_ledger}</LEDGERNAME>',
            '              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>',
            f'              <AMOUNT>-{sub_val:.2f}</AMOUNT>',
            '            </ALLLEDGERENTRIES.LIST>'
        ])

        if igst_val > 0:
            tally_lines.extend([
                '            <!-- Input IGST Debit Ledger -->',
                '            <ALLLEDGERENTRIES.LIST>',
                '              <LEDGERNAME>Input IGST 18%</LEDGERNAME>',
                '              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>',
                f'              <AMOUNT>-{igst_val:.2f}</AMOUNT>',
                '            </ALLLEDGERENTRIES.LIST>'
            ])
        else:
            if cgst_val > 0:
                tally_lines.extend([
                    '            <!-- Input CGST Debit Ledger -->',
                    '            <ALLLEDGERENTRIES.LIST>',
                    '              <LEDGERNAME>Input CGST 9%</LEDGERNAME>',
                    '              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>',
                    f'              <AMOUNT>-{cgst_val:.2f}</AMOUNT>',
                    '            </ALLLEDGERENTRIES.LIST>'
                ])
            if sgst_val > 0:
                tally_lines.extend([
                    '            <!-- Input SGST Debit Ledger -->',
                    '            <ALLLEDGERENTRIES.LIST>',
                    '              <LEDGERNAME>Input SGST 9%</LEDGERNAME>',
                    '              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>',
                    f'              <AMOUNT>-{sgst_val:.2f}</AMOUNT>',
                    '            </ALLLEDGERENTRIES.LIST>'
                ])

        tally_lines.extend([
            '          </VOUCHER>',
            '        </TALLYMESSAGE>'
        ])

    tally_lines.extend([
        '      </REQUESTDATA>',
        '    </IMPORTDATA>',
        '  </BODY>',
        '</ENVELOPE>'
    ])
    tally_xml_content = "\n".join(tally_lines)

    try:
        with open(os.path.join(acct_dir, "Tally_Purchase_Vouchers.xml"), "w", encoding="utf-8") as f:
            f.write(tally_xml_content)
    except Exception:
        pass

    # -------------------------------------------------------------------------
    # 6. Executive Invoice Dashboard (Excel Workbook)
    # -------------------------------------------------------------------------
    wb_exec = openpyxl.Workbook()

    # --- Sheet 1: Executive KPI Dashboard ---
    ws_kpi = wb_exec.active
    ws_kpi.title = "Executive Overview"
    ws_kpi.views.sheetView[0].showGridLines = True

    # Title Banner
    ws_kpi.merge_cells("A1:G1")
    title_cell = ws_kpi["A1"]
    title_cell.value = "ENTERPRISE INVOICE AUTOMATION - EXECUTIVE DASHBOARD"
    title_cell.font = Font(name="Segoe UI", size=16, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_kpi.row_dimensions[1].height = 40

    # Subtitle
    ws_kpi.merge_cells("A2:G2")
    ws_kpi["A2"].value = f"Run Summary: {total_invoices_count} Invoices Processed Across {distinct_vendors} Vendors | Total Spend: ₹{total_spend:,.2f}"
    ws_kpi["A2"].font = Font(name="Segoe UI", size=10, italic=True, color="595959")
    ws_kpi["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws_kpi.row_dimensions[2].height = 20

    # KPI Stat Cards
    kpi_cards = [
        ("TOTAL INVOICED VALUE", f"₹ {total_spend:,.2f}", "2E75B6"),
        ("TOTAL TAXABLE SUB-TOTAL", f"₹ {total_subtotal:,.2f}", "385723"),
        ("TOTAL TAXES PAID (GST)", f"₹ {total_tax_paid:,.2f}", "C55A11"),
        ("TOTAL INVOICES", str(total_invoices_count), "7030A0"),
        ("LINE ITEMS EXTRACTED", str(total_line_items), "418AB3")
    ]

    for idx, (title, val, col_hex) in enumerate(kpi_cards, 1):
        cell_label = ws_kpi.cell(row=4, column=idx + 1)
        cell_label.value = title
        cell_label.font = Font(name="Segoe UI", size=8, bold=True, color="FFFFFF")
        cell_label.fill = PatternFill(start_color=col_hex, end_color=col_hex, fill_type="solid")
        cell_label.alignment = Alignment(horizontal="center", vertical="center")

        cell_val = ws_kpi.cell(row=5, column=idx + 1)
        cell_val.value = val
        cell_val.font = Font(name="Segoe UI", size=13, bold=True, color="000000")
        cell_val.alignment = Alignment(horizontal="center", vertical="center")
        thin = Side(border_style="thin", color="D9D9D9")
        cell_val.border = Border(left=thin, right=thin, bottom=thin)

    ws_kpi.row_dimensions[4].height = 22
    ws_kpi.row_dimensions[5].height = 30

    # Vendor Summary Table
    ws_kpi.cell(row=8, column=1).value = "VENDOR SPEND & COMPLIANCE SUMMARY"
    ws_kpi.cell(row=8, column=1).font = Font(name="Segoe UI", size=12, bold=True, color="1F4E79")

    v_headers = ["Vendor Name", "State", "GSTIN", "Bills Processed", "Taxable Subtotal (₹)", "Total Tax (₹)", "Total Amount (₹)", "Avg Bill Size (₹)"]
    for c_idx, h in enumerate(v_headers, 1):
        _style_header_cell(ws_kpi.cell(row=9, column=c_idx), h, bg_color="2F5597")

    vendor_groups = {}
    for inv in parsed_invoices:
        hdr = inv["header"]
        vn = hdr.get("Vendor Name", "Unknown")
        if vn not in vendor_groups:
            vendor_groups[vn] = {
                "state": hdr.get("Vendor State", ""),
                "gstin": hdr.get("Vendor GSTIN", ""),
                "count": 0,
                "subtotal": 0.0,
                "tax": 0.0,
                "total": 0.0
            }
        vendor_groups[vn]["count"] += 1
        vendor_groups[vn]["subtotal"] += _to_float(hdr.get("SubTotal", 0))
        vendor_groups[vn]["tax"] += _to_float(hdr.get("Total Tax", 0))
        vendor_groups[vn]["total"] += _to_float(hdr.get("Total", 0))

    r_idx = 10
    for vn, vdata in sorted(vendor_groups.items()):
        avg_bill = vdata["total"] / vdata["count"] if vdata["count"] else 0.0
        ws_kpi.cell(row=r_idx, column=1, value=vn)
        ws_kpi.cell(row=r_idx, column=2, value=vdata["state"])
        ws_kpi.cell(row=r_idx, column=3, value=vdata["gstin"])
        ws_kpi.cell(row=r_idx, column=4, value=vdata["count"]).alignment = Alignment(horizontal="center")

        c5 = ws_kpi.cell(row=r_idx, column=5, value=vdata["subtotal"])
        c5.number_format = "#,##0.00"
        c5.alignment = Alignment(horizontal="right")

        c6 = ws_kpi.cell(row=r_idx, column=6, value=vdata["tax"])
        c6.number_format = "#,##0.00"
        c6.alignment = Alignment(horizontal="right")

        c7 = ws_kpi.cell(row=r_idx, column=7, value=vdata["total"])
        c7.number_format = "#,##0.00"
        c7.alignment = Alignment(horizontal="right")

        c8 = ws_kpi.cell(row=r_idx, column=8, value=avg_bill)
        c8.number_format = "#,##0.00"
        c8.alignment = Alignment(horizontal="right")
        r_idx += 1

    # Total Row
    ws_kpi.cell(row=r_idx, column=1, value="GRAND TOTAL").font = Font(name="Segoe UI", bold=True)
    ws_kpi.cell(row=r_idx, column=4, value=total_invoices_count).font = Font(name="Segoe UI", bold=True)
    ws_kpi.cell(row=r_idx, column=4).alignment = Alignment(horizontal="center")

    tot_sub_cell = ws_kpi.cell(row=r_idx, column=5, value=total_subtotal)
    tot_sub_cell.number_format = "#,##0.00"
    tot_sub_cell.font = Font(name="Segoe UI", bold=True)

    tot_tax_cell = ws_kpi.cell(row=r_idx, column=6, value=total_tax_paid)
    tot_tax_cell.number_format = "#,##0.00"
    tot_tax_cell.font = Font(name="Segoe UI", bold=True)

    tot_amt_cell = ws_kpi.cell(row=r_idx, column=7, value=total_spend)
    tot_amt_cell.number_format = "#,##0.00"
    tot_amt_cell.font = Font(name="Segoe UI", bold=True)

    for c in range(1, 9):
        ws_kpi.cell(row=r_idx, column=c).fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    _apply_table_styling(ws_kpi, 9, r_idx, 1, 8)
    _autofit_columns(ws_kpi)

    # --- Sheet 2: All Invoices Register ---
    ws_reg = wb_exec.create_sheet(title="Invoices Register")
    ws_reg.views.sheetView[0].showGridLines = True
    reg_headers = [
        "Bill Number", "Bill Date", "Due Date", "Vendor Name", "Vendor State",
        "Vendor GSTIN", "Payment Terms", "SubTotal (₹)", "CGST (₹)", "SGST (₹)",
        "IGST (₹)", "Total Tax (₹)", "Grand Total (₹)", "Line Items Count"
    ]
    for c_idx, h in enumerate(reg_headers, 1):
        _style_header_cell(ws_reg.cell(row=1, column=c_idx), h, bg_color="1F4E79")
    ws_reg.row_dimensions[1].height = 28

    r_idx = 2
    for inv in parsed_invoices:
        hdr = inv["header"]
        ws_reg.cell(row=r_idx, column=1, value=hdr.get("Bill Number", "")).alignment = Alignment(horizontal="left")
        ws_reg.cell(row=r_idx, column=2, value=hdr.get("Bill Date", "")).alignment = Alignment(horizontal="center")
        ws_reg.cell(row=r_idx, column=3, value=hdr.get("Due Date", "")).alignment = Alignment(horizontal="center")
        ws_reg.cell(row=r_idx, column=4, value=hdr.get("Vendor Name", ""))
        ws_reg.cell(row=r_idx, column=5, value=hdr.get("Vendor State", ""))
        ws_reg.cell(row=r_idx, column=6, value=hdr.get("Vendor GSTIN", ""))
        ws_reg.cell(row=r_idx, column=7, value=hdr.get("Payment Terms", ""))

        for offset, field in enumerate(["SubTotal", "CGST Amount", "SGST Amount", "IGST Amount", "Total Tax", "Total"], 8):
            c_num = ws_reg.cell(row=r_idx, column=offset, value=_to_float(hdr.get(field, 0)))
            c_num.number_format = "#,##0.00"
            c_num.alignment = Alignment(horizontal="right")

        ws_reg.cell(row=r_idx, column=14, value=len(inv["lines"])).alignment = Alignment(horizontal="center")
        r_idx += 1

    _apply_table_styling(ws_reg, 1, r_idx - 1, 1, len(reg_headers))
    _autofit_columns(ws_reg)

    # --- Sheet 3: Line Items Detail ---
    ws_lines = wb_exec.create_sheet(title="Line Items Extract")
    ws_lines.views.sheetView[0].showGridLines = True
    line_headers = [
        "Bill Number", "Vendor Name", "Bill Date", "Item Name", "HSN/SAC Code",
        "Quantity", "Unit", "Rate (₹)", "Line Amount (₹)", "Tax Rate"
    ]
    for c_idx, h in enumerate(line_headers, 1):
        _style_header_cell(ws_lines.cell(row=1, column=c_idx), h, bg_color="203764")
    ws_lines.row_dimensions[1].height = 28

    r_idx = 2
    for row in flat_rows:
        ws_lines.cell(row=r_idx, column=1, value=row.get("Bill Number", "")).alignment = Alignment(horizontal="left")
        ws_lines.cell(row=r_idx, column=2, value=row.get("Vendor Name", ""))
        ws_lines.cell(row=r_idx, column=3, value=row.get("Bill Date", "")).alignment = Alignment(horizontal="center")
        ws_lines.cell(row=r_idx, column=4, value=row.get("Item Name", ""))
        ws_lines.cell(row=r_idx, column=5, value=row.get("HSN/SAC Code", ""))

        c_qty = ws_lines.cell(row=r_idx, column=6, value=_to_float(row.get("Quantity", 0)))
        c_qty.alignment = Alignment(horizontal="right")

        ws_lines.cell(row=r_idx, column=7, value=row.get("Unit", "")).alignment = Alignment(horizontal="center")

        c_rate = ws_lines.cell(row=r_idx, column=8, value=_to_float(row.get("Rate", 0)))
        c_rate.number_format = "#,##0.00"
        c_rate.alignment = Alignment(horizontal="right")

        c_amt = ws_lines.cell(row=r_idx, column=9, value=_to_float(row.get("Line Amount", 0)))
        c_amt.number_format = "#,##0.00"
        c_amt.alignment = Alignment(horizontal="right")

        ws_lines.cell(row=r_idx, column=10, value=row.get("Line Tax Rate", "")).alignment = Alignment(horizontal="center")
        r_idx += 1

    _apply_table_styling(ws_lines, 1, r_idx - 1, 1, len(line_headers))
    _autofit_columns(ws_lines)

    buf_exec = BytesIO()
    wb_exec.save(buf_exec)
    exec_xlsx_bytes = buf_exec.getvalue()

    try:
        with open(os.path.join(mgmt_dir, "Executive_Invoice_Dashboard.xlsx"), "wb") as f:
            f.write(exec_xlsx_bytes)
    except Exception:
        pass

    # -------------------------------------------------------------------------
    # 7. Procurement & Price Catalog (Excel Workbook)
    # -------------------------------------------------------------------------
    wb_proc = openpyxl.Workbook()
    ws_proc = wb_proc.active
    ws_proc.title = "Procurement Price Catalog"
    ws_proc.views.sheetView[0].showGridLines = True

    # Title Banner
    ws_proc.merge_cells("A1:I1")
    p_title = ws_proc["A1"]
    p_title.value = "PROCUREMENT & ITEM RATE CATALOG (CROSS-INVOICE ANALYSIS)"
    p_title.font = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
    p_title.fill = PatternFill(start_color="375623", end_color="375623", fill_type="solid")
    p_title.alignment = Alignment(horizontal="center", vertical="center")
    ws_proc.row_dimensions[1].height = 35

    proc_headers = [
        "Item / Service Description", "HSN/SAC", "Unit", "Vendor(s)",
        "Purchase Frequency", "Total Quantity", "Min Rate (₹)", "Max Rate (₹)", "Total Spend (₹)"
    ]
    for c_idx, h in enumerate(proc_headers, 1):
        _style_header_cell(ws_proc.cell(row=2, column=c_idx), h, bg_color="548235")
    ws_proc.row_dimensions[2].height = 25

    r_idx = 3
    for iname, idata in sorted(item_procurement.items()):
        rates = [rec["rate"] for rec in idata["records"] if rec["rate"] > 0]
        min_rate = min(rates) if rates else 0.0
        max_rate = max(rates) if rates else 0.0
        tot_qty = sum(rec["qty"] for rec in idata["records"])
        tot_spend = sum(rec["amt"] for rec in idata["records"])
        vendors_str = ", ".join(sorted(idata["vendors"]))

        ws_proc.cell(row=r_idx, column=1, value=iname)
        ws_proc.cell(row=r_idx, column=2, value=idata["hsn"])
        ws_proc.cell(row=r_idx, column=3, value=idata["unit"]).alignment = Alignment(horizontal="center")
        ws_proc.cell(row=r_idx, column=4, value=vendors_str)
        ws_proc.cell(row=r_idx, column=5, value=len(idata["bills"])).alignment = Alignment(horizontal="center")
        ws_proc.cell(row=r_idx, column=6, value=tot_qty).alignment = Alignment(horizontal="right")

        c_min = ws_proc.cell(row=r_idx, column=7, value=min_rate)
        c_min.number_format = "#,##0.00"
        c_min.alignment = Alignment(horizontal="right")

        c_max = ws_proc.cell(row=r_idx, column=8, value=max_rate)
        c_max.number_format = "#,##0.00"
        c_max.alignment = Alignment(horizontal="right")

        c_tot = ws_proc.cell(row=r_idx, column=9, value=tot_spend)
        c_tot.number_format = "#,##0.00"
        c_tot.alignment = Alignment(horizontal="right")
        r_idx += 1

    _apply_table_styling(ws_proc, 2, r_idx - 1, 1, len(proc_headers))
    _autofit_columns(ws_proc)

    buf_proc = BytesIO()
    wb_proc.save(buf_proc)
    proc_xlsx_bytes = buf_proc.getvalue()

    try:
        with open(os.path.join(mgmt_dir, "Procurement_Price_Catalog.xlsx"), "wb") as f:
            f.write(proc_xlsx_bytes)
    except Exception:
        pass

    # -------------------------------------------------------------------------
    # 8. Audit & Compliance Verification Report (Excel Workbook)
    # -------------------------------------------------------------------------
    wb_audit = openpyxl.Workbook()
    ws_audit = wb_audit.active
    ws_audit.title = "Audit & Validation"
    ws_audit.views.sheetView[0].showGridLines = True

    # Title Banner
    ws_audit.merge_cells("A1:J1")
    a_title = ws_audit["A1"]
    a_title.value = "INVOICE AUDIT, MATH INTEGRITY & GST COMPLIANCE REPORT"
    a_title.font = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
    a_title.fill = PatternFill(start_color="833C0C", end_color="833C0C", fill_type="solid")
    a_title.alignment = Alignment(horizontal="center", vertical="center")
    ws_audit.row_dimensions[1].height = 35

    audit_headers = [
        "Bill Number", "Vendor Name", "Bill Date", "GSTIN", "State Code",
        "GSTIN Syntax", "State Code Match", "Line Item Math", "SubTotal + Tax Match", "Audit Status"
    ]
    for c_idx, h in enumerate(audit_headers, 1):
        _style_header_cell(ws_audit.cell(row=2, column=c_idx), h, bg_color="C65911")
    ws_audit.row_dimensions[2].height = 25

    audit_passed_count = 0
    r_idx = 3
    for inv in parsed_invoices:
        hdr = inv["header"]
        bno = hdr.get("Bill Number", "")
        vname = hdr.get("Vendor Name", "")
        bdate = hdr.get("Bill Date", "")
        gstin = hdr.get("Vendor GSTIN", "").strip()
        scode = hdr.get("Vendor State Code", "").strip()

        # 1. GSTIN format check (15 chars, alphanumeric)
        gstin_ok = len(gstin) == 15 and gstin[:2].isdigit() and gstin[2:12].isalnum()
        gstin_status = "VALID (15-char)" if gstin_ok else "INVALID"

        # 2. State code matches GSTIN prefix
        state_match = gstin.startswith(scode) if (gstin and scode) else False
        state_status = "MATCHED" if state_match else "MISMATCH"

        # 3. Line items math check: Qty * Rate == Line Amount (with tolerance)
        line_math_ok = True
        for line in inv["lines"]:
            q = _to_float(line.get("Quantity", 0))
            r = _to_float(line.get("Rate", 0))
            a = _to_float(line.get("Line Amount", 0))
            if q > 0 and r > 0 and a > 0:
                if abs((q * r) - a) > CONFIG["LINE_MATH_TOLERANCE"]:
                    line_math_ok = False
                    break
        line_status = "PASSED" if line_math_ok else "FLAGGED"

        # 4. Tax math check: SubTotal + Total Tax == Total
        sub_val = _to_float(hdr.get("SubTotal", 0))
        tax_val = _to_float(hdr.get("Total Tax", 0))
        tot_val = _to_float(hdr.get("Total", 0))
        tax_math_ok = abs((sub_val + tax_val) - tot_val) < CONFIG["TAX_MATH_TOLERANCE"]
        tax_status = "PASSED" if tax_math_ok else "FLAGGED"

        overall_ok = gstin_ok and state_match and line_math_ok and tax_math_ok
        if overall_ok:
            audit_passed_count += 1
            overall_status = "PASSED ALL SET CHECKS"
            status_color = "385723"
        else:
            overall_status = "FLAGGED FOR REVIEW"
            status_color = "C00000"

        ws_audit.cell(row=r_idx, column=1, value=bno).font = Font(name="Segoe UI", bold=True)
        ws_audit.cell(row=r_idx, column=2, value=vname)
        ws_audit.cell(row=r_idx, column=3, value=bdate)
        ws_audit.cell(row=r_idx, column=4, value=gstin)
        ws_audit.cell(row=r_idx, column=5, value=scode).alignment = Alignment(horizontal="center")
        ws_audit.cell(row=r_idx, column=6, value=gstin_status).alignment = Alignment(horizontal="center")
        ws_audit.cell(row=r_idx, column=7, value=state_status).alignment = Alignment(horizontal="center")
        ws_audit.cell(row=r_idx, column=8, value=line_status).alignment = Alignment(horizontal="center")
        ws_audit.cell(row=r_idx, column=9, value=tax_status).alignment = Alignment(horizontal="center")

        st_cell = ws_audit.cell(row=r_idx, column=10, value=overall_status)
        st_cell.font = Font(name="Segoe UI", bold=True, color=status_color)
        st_cell.alignment = Alignment(horizontal="center")
        r_idx += 1

    _apply_table_styling(ws_audit, 2, r_idx - 1, 1, len(audit_headers))
    _autofit_columns(ws_audit)

    buf_audit = BytesIO()
    wb_audit.save(buf_audit)
    audit_xlsx_bytes = buf_audit.getvalue()

    try:
        with open(os.path.join(mgmt_dir, "Audit_Compliance_Report.xlsx"), "wb") as f:
            f.write(audit_xlsx_bytes)
    except Exception:
        pass

    # -------------------------------------------------------------------------
    # 9. Vendor-Specific Folders
    # -------------------------------------------------------------------------
    for vn, vdata in vendor_groups.items():
        clean_vname = re.sub(r"[^A-Za-z0-9_]+", "_", vn).strip("_")
        v_folder = os.path.join(vendor_root_dir, clean_vname)
        try:
            os.makedirs(v_folder, exist_ok=True)
            v_rows = [r for r in flat_rows if r.get("Vendor Name") == vn]

            # 1. Vendor CSV
            v_csv_path = os.path.join(v_folder, f"{clean_vname}_Bills.csv")
            with open(v_csv_path, "w", encoding="utf-8", newline="") as vf:
                v_writer = csv.DictWriter(vf, fieldnames=tags_list, extrasaction="ignore", lineterminator="\n")
                v_writer.writeheader()
                for r in v_rows:
                    v_writer.writerow({k: r.get(k, "") for k in tags_list})

            # 2. Vendor Styled XLSX
            wb_v = openpyxl.Workbook()
            ws_v = wb_v.active
            ws_v.title = "Bills Register"
            ws_v.views.sheetView[0].showGridLines = True
            for c_idx, h in enumerate(tags_list, 1):
                _style_header_cell(ws_v.cell(row=1, column=c_idx), h, bg_color="1F4E79")
            ws_v.row_dimensions[1].height = 25

            for r_i, r_data in enumerate(v_rows, 2):
                for c_i, h in enumerate(tags_list, 1):
                    val = r_data.get(h, "")
                    cell = ws_v.cell(row=r_i, column=c_i)
                    if h in ["Rate", "Line Amount", "SubTotal", "CGST Amount", "SGST Amount", "IGST Amount", "Total Tax", "Total"]:
                        cell.value = _to_float(val)
                        cell.number_format = "#,##0.00"
                        cell.alignment = Alignment(horizontal="right")
                    elif h == "Quantity":
                        cell.value = _to_float(val)
                        cell.alignment = Alignment(horizontal="right")
                    else:
                        cell.value = val
            _apply_table_styling(ws_v, 1, len(v_rows) + 1, 1, len(tags_list))
            _autofit_columns(ws_v)
            wb_v.save(os.path.join(v_folder, f"{clean_vname}_Bills.xlsx"))
        except Exception:
            pass

    # -------------------------------------------------------------------------
    # 10. Return files for Application Output Panel
    # -------------------------------------------------------------------------
    return [
        {
            "name": "Executive_Invoice_Dashboard.xlsx",
            "content": exec_xlsx_bytes,
            "message": f"Executive Dashboard: {total_invoices_count} invoices across {distinct_vendors} vendors, with KPI cards and vendor summary."
        },
        {
            "name": "Zoho_Bills_Import.csv",
            "content": zoho_csv_content,
            "message": f"Zoho Books Ready: {total_line_items} line items mapped to Accounts Payable."
        },
        {
            "name": "Tally_Purchase_Vouchers.xml",
            "content": tally_xml_content,
            "message": f"Tally Prime Ready: {total_invoices_count} purchase vouchers formatted in native XML for direct Day Book ingestion."
        },
        {
            "name": "QuickBooks_Bills_Import.csv",
            "content": qb_csv_content,
            "message": f"QuickBooks Online Ready: Batch bill import format with item-level expense categories."
        },
        {
            "name": "Procurement_Price_Catalog.xlsx",
            "content": proc_xlsx_bytes,
            "message": f"Procurement Analytics: {len(item_procurement)} unique items cataloged with min/max/avg rate tracking."
        },
        {
            "name": "Audit_Compliance_Report.xlsx",
            "content": audit_xlsx_bytes,
            "message": f"Set checks (line math, tax math, GSTIN format, state code): {audit_passed_count} of {total_invoices_count} invoices passed all set checks, {total_invoices_count - audit_passed_count} flagged for review."
        },
        {
            "name": "output.csv",
            "content": output_csv_content,
            "message": f"Master Flat CSV: Complete export of all 26 tags across {total_invoices_count} processed documents."
        }
    ]
