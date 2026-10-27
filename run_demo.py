#!/usr/bin/env python3
"""
Workflow Automation Demo Runner
Reads sample PDF invoices, runs template detection and field extraction,
verifies invoices against defined set checks, and generates Excel registers,
audit summaries, and accounting import files.

Built by Gentle Reminder (https://gentlereminder.in/)
"""

import os
import sys
import glob
import argparse
import importlib.util
from datetime import datetime

# ==============================================================================
# CONFIGURATION BLOCK
# Change default directories and detection parameters here.
# ==============================================================================
CONFIG = {
    # Default folder containing sample invoice PDFs
    "INPUT_DIR": "sample-invoices",

    # Default folder where generated reports and import files will land
    "OUTPUT_DIR": "sample-output",

    # Minimum score (0-100) required to match a vendor template
    "DETECTION_THRESHOLD": 70,

    # Folder containing vendor extraction templates
    "TEMPLATES_DIR": "templates",
}


def load_templates(templates_root):
    """Dynamically loads detect and extract modules for all vendor templates."""
    templates = {}
    if not os.path.isdir(templates_root):
        raise FileNotFoundError(f"Templates directory not found: {templates_root}")

    for vendor_dir in sorted(os.listdir(templates_root)):
        vpath = os.path.join(templates_root, vendor_dir)
        if not os.path.isdir(vpath):
            continue

        detect_file = os.path.join(vpath, "detect.py")
        extract_file = os.path.join(vpath, "extract.py")

        if os.path.isfile(detect_file) and os.path.isfile(extract_file):
            # Load detect module
            spec_d = importlib.util.spec_from_file_location(f"detect_{vendor_dir}", detect_file)
            mod_d = importlib.util.module_from_spec(spec_d)
            spec_d.loader.exec_module(mod_d)

            # Load extract module
            spec_e = importlib.util.spec_from_file_location(f"extract_{vendor_dir}", extract_file)
            mod_e = importlib.util.module_from_spec(spec_e)
            spec_e.loader.exec_module(mod_e)

            templates[vendor_dir] = {
                "detect": mod_d.detect,
                "extract": mod_e.extract,
            }

    return templates


def parse_pdf(pdf_path):
    """
    Parses a PDF into structured text pages and layout lines using
    pypdfium2 and pdfplumber.
    """
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
        "pypdfium2": {
            "page_count": len(page_texts),
            "pages": [{"index": i, "text": t} for i, t in enumerate(page_texts)],
        },
        "pdfplumber": {
            "granularities": {
                "line": plumber_lines,
            }
        },
    }


def main():
    parser = argparse.ArgumentParser(
        description="Demo: read invoices from PDF, check them against set checks, and write Excel and Tally-ready import files."
    )
    parser.add_argument(
        "--input",
        "-i",
        default=CONFIG["INPUT_DIR"],
        help=f"Directory containing PDF invoices (default: {CONFIG['INPUT_DIR']})",
    )
    parser.add_argument(
        "--output",
        "-o",
        default=CONFIG["OUTPUT_DIR"],
        help=f"Directory to write output reports and import files (default: {CONFIG['OUTPUT_DIR']})",
    )
    parser.add_argument(
        "--threshold",
        "-t",
        type=int,
        default=CONFIG["DETECTION_THRESHOLD"],
        help=f"Minimum detection score (default: {CONFIG['DETECTION_THRESHOLD']})",
    )
    args = parser.parse_args()

    base_dir = os.path.dirname(os.path.abspath(__file__))
    input_dir = args.input if os.path.isabs(args.input) else os.path.join(base_dir, args.input)
    output_dir = args.output if os.path.isabs(args.output) else os.path.join(base_dir, args.output)
    templates_dir = os.path.join(base_dir, CONFIG["TEMPLATES_DIR"])

    print("=" * 70)
    print("PDF INVOICE TO EXCEL & TALLY-READY WORKFLOW DEMO")
    print("=" * 70)
    print(f"Input Directory  : {input_dir}")
    print(f"Output Directory : {output_dir}")
    print(f"Templates        : {templates_dir}")
    print("-" * 70)

    # 1. Load templates
    templates = load_templates(templates_dir)
    print(f"Loaded {len(templates)} vendor templates: {', '.join(templates.keys())}")

    # 2. Find PDF files
    pdf_files = sorted(glob.glob(os.path.join(input_dir, "*.pdf")))
    if not pdf_files:
        print(f"Error: No PDF files found in {input_dir}")
        sys.exit(1)

    print(f"Discovered {len(pdf_files)} invoice PDF(s) to process.\n")

    # 3. Process documents
    processed_docs = []
    template_counts = {k: 0 for k in templates}
    unmatched_files = []

    for idx, pdf_path in enumerate(pdf_files, 1):
        filename = os.path.basename(pdf_path)
        print(f"[{idx:02d}/{len(pdf_files):02d}] Reading {filename}...", end=" ", flush=True)

        parsed_context = parse_pdf(pdf_path)

        # Detect matching template
        best_match = None
        best_score = 0
        for vendor_name, tmpl in templates.items():
            score = tmpl["detect"](parsed_context, pdf_path)
            if score >= args.threshold and score > best_score:
                best_score = score
                best_match = vendor_name

        if best_match:
            template_counts[best_match] += 1
            extracted_data = templates[best_match]["extract"](parsed_context, pdf_path)
            bill_num = extracted_data.get("Bill Number") or "N/A"
            vname = extracted_data.get("Vendor Name") or best_match
            print(f"Matched {best_match} (score {best_score}) -> Bill #{bill_num}")

            processed_docs.append({
                "file": filename,
                "template": best_match,
                "extraction": extracted_data,
            })
        else:
            print("UNMATCHED")
            unmatched_files.append(filename)

    print("\n" + "-" * 70)
    print("EXTRACTION SUMMARY:")
    for vname, cnt in template_counts.items():
        print(f"  • {vname:18s}: {cnt} invoice(s)")
    if unmatched_files:
        print(f"  • Unmatched: {len(unmatched_files)} file(s): {', '.join(unmatched_files)}")

    # 4. Generate outputs and reports
    print("\nGenerating consolidated reports, verification audit, and import files...")
    import output as out_mod

    result_files = out_mod.output(processed_docs, target_dir=output_dir)

    print("\n" + "=" * 70)
    print("RUN COMPLETE - OUTPUTS GENERATED:")
    print("=" * 70)
    for rf in result_files:
        print(f"  [OK] {rf['name']}")
        print(f"       -> {rf['message']}")

    print("\nAll files written to:")
    print(f"  {output_dir}")
    print("=" * 70)


if __name__ == "__main__":
    main()
