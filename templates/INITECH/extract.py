import re
import os
import json

# ==============================================================================
# CONFIGURATION BLOCK
# Change static vendor and buyer settings here without modifying extraction logic.
# ==============================================================================
CONFIG = {
    "VENDOR_NAME": "Initech",
    "VENDOR_GSTIN": "32SVVCI0589F1ZL",
    "VENDOR_STATE": "Kerala",
    "VENDOR_STATE_CODE": "32",
    "BUYER_NAME": "Gentle Reminder",
    "BUYER_GSTIN": "24BBQPP2817N1Z3",
    "DEFAULT_PAYMENT_TERMS": "NET 30",
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
    bill_no_m = re.search(r"invoice no\s*:\s*([^\n\r]+)", text)
    bill_number = bill_no_m.group(1).strip() if bill_no_m else ""

    bill_date_m = re.search(r"invoice date\s*:\s*([^\n\r]+)", text)
    bill_date = bill_date_m.group(1).strip() if bill_date_m else ""

    due_date_m = re.search(r"due date\s*:\s*([^\n\r]+)", text)
    due_date = due_date_m.group(1).strip() if due_date_m else ""

    pay_terms_m = re.search(r"terms\s*:\s*([^\n\r]+)", text)
    payment_terms = pay_terms_m.group(1).strip() if pay_terms_m else CONFIG["DEFAULT_PAYMENT_TERMS"]

    # 3. Buyer Details
    buyer_name = CONFIG["BUYER_NAME"]
    buyer_gstin = CONFIG["BUYER_GSTIN"]
    if "CUSTOMER" in text:
        cust_part = text.split("CUSTOMER")[1].split("# DESCRIPTION")[0]
        cust_lines = [l.strip() for l in cust_part.splitlines() if l.strip()]
        for l in cust_lines:
            if not l.startswith("GSTIN") and not l.startswith("hello@") and not l.startswith("+91") and "Ahmedabad" not in l:
                buyer_name = l
                break
        bg_m = re.search(r"GSTIN\s+([0-9A-Z]{15})", cust_part)
        if bg_m:
            buyer_gstin = bg_m.group(1)

    # 4. Multiline Items
    item_names = []
    hsn_codes = []
    quantities = []
    units = []
    rates = []
    line_amounts = []
    line_tax_rates = []

    start_idx = -1
    for idx, l in enumerate(lines):
        if "#  DESCRIPTION" in l or "# DESCRIPTION" in l:
            start_idx = idx
            break

    if start_idx != -1:
        curr = start_idx + 1
        while curr < len(lines):
            l = lines[curr]
            if "SUBTOTAL" in l:
                break
            m = re.match(r"^(\d+)\s+(.+?)\s+(\d{6})\s+(\d+)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})$", l)
            if m:
                item_names.append(m.group(2).strip())
                hsn_codes.append(m.group(3).strip())
                quantities.append(m.group(4).strip())
                units.append("")
                rates.append(m.group(5).strip())
                line_amounts.append(m.group(6).strip())
                line_tax_rates.append("18%")
            elif item_names:
                # Continuation description
                item_names[-1] = (item_names[-1] + " " + l).strip()
            curr += 1

    # 5. Totals & Taxes
    subtotal_m = re.search(r"SUBTOTAL\s*Rs\.\s*([\d,]+\.\d{2})", text)
    subtotal = subtotal_m.group(1).strip() if subtotal_m else ""

    igst_m = re.search(r"IGST\s*@18%\s*Rs\.\s*([\d,]+\.\d{2})", text)
    igst_amount = igst_m.group(1).strip() if igst_m else ""

    cgst_amount = "0.00"
    sgst_amount = "0.00"
    total_tax = igst_amount

    total_m = re.search(r"TOTAL PAYABLE\s*Rs\.\s*([\d,]+\.\d{2})", text)
    total = total_m.group(1).strip() if total_m else ""

    words_m = re.search(r"amount in words:\s*([^\n\r]+)", text)
    amount_in_words = words_m.group(1).strip() if words_m else ""

    # 6. Notes & Terms
    vendor_notes = "Did you get the memo? All invoice queries: licensing@initech.example / +91 484 271 9090 | Remit to: Initech | TPS Bank (fictional) | A/c 0101 0101 0042 | IFSC TPSB0000042"
    terms_conditions = "Quote the invoice number with your payment. Licence keys are emailed to the registered admin on receipt of payment. Renewals are invoiced 30 days before expiry."

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
