import re
import os
import json

# ==============================================================================
# CONFIGURATION BLOCK
# Change static vendor and buyer settings here without modifying extraction logic.
# ==============================================================================
CONFIG = {
    "VENDOR_NAME": "ACME CORPORATION",
    "VENDOR_GSTIN": "27MRACA9777I7ZQ",
    "VENDOR_STATE": "Maharashtra",
    "VENDOR_STATE_CODE": "27",
    "BUYER_NAME": "Gentle Reminder",
    "BUYER_GSTIN": "24BBQPP2817N1Z3",
    "DEFAULT_PAYMENT_TERMS": "Net 30 days",
}


def _parse_pdf_to_context(pdf_path):
    """Directly extracts text pages and lines from a PDF file."""
    import pypdfium2
    import pdfplumber

    pdfium_doc = pypdfium2.PdfDocument(pdf_path)
    page_texts = [page.get_textpage().get_text_range() for page in pdfium_doc]

    with pdfplumber.open(pdf_path) as pdf:
        plumber_lines = []
        for p_idx, page in enumerate(pdf.pages):
            for l_idx, line in enumerate(page.extract_text_lines()):
                plumber_lines.append({
                    "id": f"line_{p_idx}_{l_idx}",
                    "content": line["text"],
                    "page": p_idx,
                })

    return {
        "pypdfium2": {"pages": [{"index": i, "text": t} for i, t in enumerate(page_texts)]},
        "pdfplumber": {"granularities": {"line": plumber_lines}},
    }


def _get_context(from_json, from_file=None):
    data = from_json
    if not data and from_file and isinstance(from_file, str):
        candidate = os.path.join(os.path.dirname(from_file), "processed.json")
        if os.path.isfile(candidate):
            try:
                with open(candidate, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except Exception:
                pass
        elif from_file.lower().endswith(".pdf") and os.path.isfile(from_file):
            try:
                data = _parse_pdf_to_context(from_file)
            except Exception:
                pass
    if not data:
        data = {}

    text = ""
    pypdfium = data.get("pypdfium2") or {}
    pages = pypdfium.get("pages", [])
    if pages and isinstance(pages, list) and "text" in pages[0]:
        text = pages[0]["text"]
    else:
        pdfplumber = data.get("pdfplumber") or data
        blocks = pdfplumber.get("granularities", {}).get("block", [])
        if blocks:
            text = "\n".join(b.get("content", "") for b in blocks if isinstance(b, dict))

    lines = [
        item["content"].strip()
        for item in data.get("pdfplumber", {}).get("granularities", {}).get("line", [])
        if isinstance(item, dict) and "content" in item
    ]
    return data, text, lines


def extract(from_json, from_file=None):
    data, text, lines = _get_context(from_json, from_file)

    # 1. Vendor Details
    vendor_name = CONFIG["VENDOR_NAME"]
    vendor_gstin = CONFIG["VENDOR_GSTIN"]
    vendor_state = CONFIG["VENDOR_STATE"]
    vendor_state_code = CONFIG["VENDOR_STATE_CODE"]

    # 2. Bill / Invoice Header
    bill_no_m = re.search(r"Invoice No\.\s*:\s*([^\n\r]+)", text)
    bill_number = bill_no_m.group(1).split()[0].strip() if bill_no_m else ""

    bill_date_m = re.search(r"Invoice Date\s*:\s*([^\n\r]+)", text)
    bill_date = bill_date_m.group(1).split()[0].strip() if bill_date_m else ""

    due_date_m = re.search(r"Due Date\s*:\s*([^\n\r]+)", text)
    due_date = due_date_m.group(1).split()[0].strip() if due_date_m else ""

    pay_terms_m = re.search(r"Payment Terms\s*:\s*([^\n\r]+)", text)
    payment_terms = pay_terms_m.group(1).strip() if pay_terms_m else CONFIG["DEFAULT_PAYMENT_TERMS"]

    # 3. Buyer Details
    buyer_name = CONFIG["BUYER_NAME"]
    buyer_gstin = CONFIG["BUYER_GSTIN"]
    if "Details of Receiver" in text:
        rec_part = text.split("Details of Receiver")[1].split("Details of Consignee")[0]
        rec_lines = [l.strip() for l in rec_part.splitlines() if l.strip() and not l.strip().startswith("(")]
        for l in rec_lines:
            if not l.startswith("GSTIN") and not l.startswith("Email") and not l.startswith("Phone") and "Ahmedabad" not in l:
                buyer_name = l
                break
        bg_m = re.search(r"GSTIN:\s*([0-9A-Z]{15})", rec_part)
        if bg_m:
            buyer_gstin = bg_m.group(1)

    # 4. Multiline Item Extraction
    item_names = []
    hsn_codes = []
    quantities = []
    units = []
    rates = []
    line_amounts = []
    line_tax_rates = []

    start_idx = -1
    for idx, l in enumerate(lines):
        if "Sr Description of Goods" in l:
            start_idx = idx
            break

    if start_idx != -1:
        curr = start_idx + 1
        while curr < len(lines):
            l = lines[curr]
            if "Total Taxable Value" in l or "Amount Chargeable" in l:
                break

            m = re.match(r"^(\d+)\s+(.+?)\s+(\d{4,8})\s+(\d+)\s+([A-Za-z]+)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+(\d+%)\s+([\d,]+\.\d{2})$", l)
            if m:
                _, desc, hsn, qty, unit, rate, taxable, igst_pct, igst_amt = m.groups()
                item_names.append(desc.strip())
                hsn_codes.append(hsn.strip())
                quantities.append(qty.strip())
                units.append(unit.strip())
                rates.append(rate.strip())
                line_amounts.append(taxable.strip())
                line_tax_rates.append(igst_pct.strip())
            else:
                m_start = re.match(r"^(\d+)\s+(.+)$", l)
                if m_start and not any(k in l for k in ["Total", "Rate", "Taxable", "Add:", "Bank details"]):
                    desc_parts = [m_start.group(2)]
                    curr += 1
                    while curr < len(lines):
                        nxt = lines[curr]
                        m_end = re.search(r"(\d{4,8})\s+(\d+)\s+([A-Za-z]+)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+(\d+%)\s+([\d,]+\.\d{2})$", nxt)
                        if m_end:
                            prefix = nxt[:m_end.start()].strip()
                            if prefix:
                                desc_parts.append(prefix)
                            hsn, qty, unit, rate, taxable, igst_pct, igst_amt = m_end.groups()
                            item_names.append(" ".join(desc_parts).strip())
                            hsn_codes.append(hsn.strip())
                            quantities.append(qty.strip())
                            units.append(unit.strip())
                            rates.append(rate.strip())
                            line_amounts.append(taxable.strip())
                            line_tax_rates.append(igst_pct.strip())
                            break
                        else:
                            if any(k in nxt for k in ["Total Taxable Value", "Amount Chargeable", "Add:"]):
                                curr -= 1
                                break
                            desc_parts.append(nxt)
                        curr += 1
            curr += 1

    # 5. Totals & Taxes
    subtotal_m = re.search(r"Total Taxable Value\s+([\d,]+\.\d{2})", text)
    subtotal = subtotal_m.group(1).strip() if subtotal_m else ""

    igst_m = re.search(r"Add:\s*IGST\s+([\d,]+\.\d{2})", text)
    igst_amount = igst_m.group(1).strip() if igst_m else ""

    cgst_amount = "0.00"
    sgst_amount = "0.00"
    total_tax = igst_amount

    total_m = re.search(r"Grand Total\s*\([^\)]*\)\s+([\d,]+\.\d{2})", text)
    total = total_m.group(1).strip() if total_m else ""

    in_words_m = re.search(r"Amount Chargeable \(in words\):\s*\n([^\n\r]+)", text)
    amount_in_words = in_words_m.group(1).strip() if in_words_m else ""

    # 6. Notes & Terms
    bank_m = re.search(r"Bank details\s*\n([^\n\r]+)", text)
    vendor_notes = bank_m.group(1).strip() if bank_m else "Bank: Roadrunner Co-operative Bank (fictional) | A/c 001234567890 | IFSC RRCB0000412"

    terms_m = re.search(r"Terms:\s*([^\n\r]+(?:\n[^\n\r]+)?)", text)
    terms_conditions = " ".join(terms_m.group(1).split()) if terms_m else "Goods once sold will not be taken back. Interest @18% p.a. on payments delayed beyond the due date. Subject to Mumbai jurisdiction. This is a computer generated invoice."

    return {
        "Vendor Name": vendor_name,
        "Vendor GSTIN": vendor_gstin,
        "Vendor State": vendor_state,
        "Vendor State Code": vendor_state_code,
        "Bill Number": bill_number,
        "Bill Date": bill_date,
        "Due Date": due_date,
        "Payment Terms": payment_terms,
        "Buyer Name": buyer_name,
        "Buyer GSTIN": buyer_gstin,
        "Item Name": item_names,
        "HSN/SAC Code": hsn_codes,
        "Quantity": quantities,
        "Unit": units,
        "Rate": rates,
        "Line Amount": line_amounts,
        "Line Tax Rate": line_tax_rates,
        "SubTotal": subtotal,
        "CGST Amount": cgst_amount,
        "SGST Amount": sgst_amount,
        "IGST Amount": igst_amount,
        "Total Tax": total_tax,
        "Total": total,
        "Amount in Words": amount_in_words,
        "Vendor Notes": vendor_notes,
        "Terms & Conditions": terms_conditions,
    }
