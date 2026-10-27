import re
import os

# ==============================================================================
# CONFIGURATION BLOCK
# Change detection keywords and threshold score here.
# ==============================================================================
CONFIG = {
    # Match score threshold (0-100). A score >= PASS_SCORE matches this template.
    "PASS_SCORE": 70,

    # Primary identifier: Vendor GSTIN (fictional sample data)
    "VENDOR_GSTIN": "32svvci0589f1zl",

    # Secondary identifier keywords
    "VENDOR_KEYWORDS": ["initech"],
    "CONTACT_KEYWORDS": ["licensing@initech.example", "initech.example", "271 9090"],
    "ADDRESS_BANK_KEYWORDS": [
        "marine tech tower",
        "kakkanad",
        "kochi, kerala",
        "tpsb0000042",
        "tps bank",
        "did you get the memo",
        "software licensing & support"
    ],
}


def _extract_text(from_json, from_file=None):
    """Recursively extracts and normalizes all text content from from_json or from_file."""
    parts = []

    if from_json:
        def _walk(obj):
            if isinstance(obj, str):
                parts.append(obj)
            elif isinstance(obj, dict):
                for val in obj.values():
                    _walk(val)
            elif isinstance(obj, (list, tuple)):
                for item in obj:
                    _walk(item)
        _walk(from_json)

    if not parts and from_file and isinstance(from_file, str) and os.path.isfile(from_file):
        if from_file.lower().endswith(".pdf"):
            try:
                import pypdfium2
                doc = pypdfium2.PdfDocument(from_file)
                for page in doc:
                    parts.append(page.get_textpage().get_text_range())
            except Exception:
                pass
        else:
            try:
                with open(from_file, "r", errors="ignore") as fp:
                    parts.append(fp.read())
            except Exception:
                pass

    return " ".join(parts).lower()


def detect(from_json, from_file=None):
    """
    Detects whether the document matches INITECH template.
    Target Score: >= 70 for INITECH, < 70 (typically 0) for other documents.
    """
    text = _extract_text(from_json, from_file)
    if not text:
        return 0

    score = 0

    # 1. Primary Identifier: Vendor GSTIN
    if CONFIG["VENDOR_GSTIN"] in text:
        score += 50

    # 2. Vendor Name / Brand / Tagline
    if any(k in text for k in CONFIG["VENDOR_KEYWORDS"]):
        score += 35

    # 3. Vendor Contact / Domain / Phone
    if any(k in text for k in CONFIG["CONTACT_KEYWORDS"]):
        score += 15

    # 4. Vendor Address / Bank / Specific Phrases
    if any(k in text for k in CONFIG["ADDRESS_BANK_KEYWORDS"]):
        score += 15

    # 5. Vendor Invoice Number Pattern (e.g. INI-2025-04-0085 or INI- prefix)
    if re.search(r"\bini-2025-04-\d+|\bini-\d{4,}\b", text):
        score += 10

    return min(100, score)
