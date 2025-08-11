# parse_quote.py — robust extractor for auto-insurance quotes (Progressive / Travelers / generic)
# Rev: Aug 11, 2025 — Travelers OCR fixes integrated

import boto3
import re
import io
import fitz  # PyMuPDF
from PIL import Image
import traceback
from copy import deepcopy
from typing import Dict, Any, List

"""
Key behaviors aligned with product rules:
- Progressive & Travelers specific phrases supported; falls back to generic parsing.
- total_premium prefers explicit Pay-in-Full text and avoids picking up 'Your Total Savings'.
- policy_term extracted from phrases like “Total 6 month policy premium”.
- Liability: parses BI per person/accident + PD; marks selected=True only if amounts found.
- Uninsured Motorist (UM): UMBI and UMPD parsed separately.
  * If UMPD found but no deductible is visible, defaults to $250 (per spec).
  * If both UMBI and UMPD missing → selected=False and “没有选择该项目”。
- Medical Payments / Personal Injury Protection: selected=True only when keyword is present
  AND there’s a nearby numeric amount (per spec requiring both presence and amount).
- Vehicles: VIN detection with strong regex, robust model-year extraction without
  confusing addresses; extracts Collision/Comprehensive deductibles, Rental limits,
  and Roadside presence near each VIN block.
- Currency normalized with commas.

Primary entry point: extract_quote_data(file, return_raw_text=False)
Returns either dict or (dict, full_text) when return_raw_text=True.
"""

# =====================
# Regex & constants
# =====================
YEAR_RE = r"(19\d{2}|20\d{2})"
VIN_RE = r"([A-HJ-NPR-Z\d]{17})"  # excludes I, O, Q
MONEY_RE = r"\$?\d{1,3}(?:,\d{3})*(?:\.\d{2})?"
BI_PAIR_RE = r"(\$?\\d{1,3}(?:,\\d{3})*)(?:\\s*/\\s*)(\\$?\\d{1,3}(?:,\\d{3})*)"  # 30,000/60,000 etc.
# Allow limits with commas, e.g., 40/1,200
LIMIT_RE = r"\\d{1,3}(?:,\\d{3})?\\s*/\\s*\\d{1,4}(?:,\\d{3})?"

UMBI_KEYS = [
    "Uninsured/Underinsured Motorist Bodily Injury",
    "Uninsured Motorist Bodily Injury",
    "Underinsured Motorist Bodily Injury",
    "UMBI",
    # Travelers OCR variants
    "Uninsd/Underinsd Motorists",
    "Uninsd Motorists",
    "Underinsd Motorists",
]
UMPD_KEYS = [
    "Uninsured/Underinsured Motorist Property Damage",
    "Uninsured Motorist Property Damage",
    "Underinsured Motorist Property Damage",
    "UMPD",
    # Travelers OCR variants
    "Uninsd/Underinsd Motorists PD",
    "Uninsd Motorists PD",
    "Underinsd Motorists PD",
]

ADDR_STOP_WORDS = [
    "street", "st.", "st ", "road", "rd.", "rd ", "ave", "avenue", "boulevard", "blvd", "lane", "ln",
    "drive", "dr", "suite", "ste", "apt", "unit"
]

# =====================
# Entry point
# =====================

def extract_quote_data(file, return_raw_text: bool = False):
    file_bytes_raw = file.read()
    file_bytes = deepcopy(file_bytes_raw)
    pdf_bytes_for_fitz = deepcopy(file_bytes_raw)
    file_suffix = file.name.split(".")[-1].lower()

    textract = boto3.client("textract", region_name="us-east-1")

    def is_pdf_text_based(pdf_bytes):
        try:
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            for page in doc:
                if page.get_text().strip():
                    return True
            return False
        except Exception:
            return False

    try:
        if file_suffix == "pdf":
            if is_pdf_text_based(pdf_bytes_for_fitz):
                doc = fitz.open(stream=pdf_bytes_for_fitz, filetype="pdf")
                lines = []
                for page in doc:
                    lines.extend(page.get_text().splitlines())
                full_text = "\n".join(lines)
            else:
                images = pdf_to_images(file_bytes)
                if not images:
                    raise ValueError("PDF 转图片失败")
                all_text = []
                for image_data in images:
                    img = Image.open(io.BytesIO(image_data)).convert("RGB")
                    img_buffer = io.BytesIO()
                    img.save(img_buffer, format="PNG")
                    img_buffer.seek(0)
                    img_response = textract.detect_document_text(Document={"Bytes": img_buffer.read()})
                    for block in img_response.get("Blocks", []):
                        if block.get("BlockType") == "LINE":
                            all_text.append(block.get("Text", ""))
                full_text = "\n".join(all_text)
        elif file_suffix in ["jpg", "jpeg", "png"]:
            response = textract.detect_document_text(Document={"Bytes": file_bytes})
            lines = [block["Text"] for block in response["Blocks"] if block["BlockType"] == "LINE"]
            full_text = "\n".join(lines)
        else:
            raise ValueError("不支持的文件格式")

        data = {
            "company": extract_company_name(full_text),
            "total_premium": extract_total_premium(full_text),
            "policy_term": extract_policy_term(full_text),
            "liability": extract_liability(full_text),
            "uninsured_motorist": extract_uninsured_motorist(full_text),
            "medical_payment": extract_medical_payment(full_text),
            "personal_injury": extract_personal_injury(full_text),
            "vehicles": extract_vehicles(full_text),
        }

        return (data, full_text) if return_raw_text else data

    except Exception:
        traceback.print_exc()
        raise

# =====================
# Helpers
# =====================

def pdf_to_images(pdf_bytes):
    images = []
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    for page in doc:
        pix = page.get_pixmap(dpi=300, alpha=False)
        img_buffer = io.BytesIO()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img.save(img_buffer, format="PNG")
        images.append(img_buffer.getvalue())
    return images


def detect_company(text: str) -> str:
    lowered = text.lower()
    if "progressive" in lowered:
        return "Progressive"
    if "travelers" in lowered:
        return "Travelers"
    if "allstate" in lowered:
        return "Allstate"
    if "geico" in lowered:
        return "Geico"
    if "liberty mutual" in lowered:
        return "Liberty Mutual"
    if "safeco" in lowered:
        return "Safeco"
    if "state farm" in lowered:
        return "State Farm"
    if "nationwide" in lowered:
        return "Nationwide"
    return "某保险公司"


def extract_company_name(text: str) -> str:
    return detect_company(text)


def normalize_money(val: str) -> str:
    digits = re.sub(r"[^\d.]", "", val)
    if digits == "":
        return ""
    if "." in digits:
        amount = float(digits)
        return f"${amount:,.2f}"
    else:
        amount = int(digits)
        return f"${amount:,.0f}"


def extract_total_premium(text: str) -> str:
    # Prefer explicit pay-in-full amounts; avoid picking up savings lines
    patterns = [
        r"estimated\s+pay-?in-?full[\s\S]*?\$([\d,]+\.\d{2})",
        r"pay-?in-?full[\s\S]*?\$([\d,]+\.\d{2})",
        r"Total\s+\d+\s+month\s+policy\s+premium[\s\S]*?\$([\d,]+\.\d{2})",
        r"your\s+estimated\s+total\s+premium[\s\S]*?\$([\d,]+\.\d{2})",
        r"Total\s+policy\s+premium[\s\S]*?\$([\d,]+\.\d{2})",
    ]
    for p in patterns:
        m = re.search(p, text, flags=re.IGNORECASE)
        if m:
            span_start = m.start()
            guard_window = text[max(0, span_start-40):span_start+40].lower()
            if "savings" in guard_window:
                continue
            return normalize_money(m.group(1))
    # Fallback: choose the largest dollar amount >= 1,000 that is not adjacent to 'savings'
    amounts = [(m.group(0), m.start()) for m in re.finditer(MONEY_RE, text)]
    candidates = []
    for raw, pos in amounts:
        try:
            if text[max(0, pos-30):pos+30].lower().find("savings") != -1:
                continue
            val = float(re.sub(r"[,$]", "", raw))
            if val >= 1000:
                candidates.append(val)
        except Exception:
            continue
    if candidates:
        return f"${max(candidates):,.2f}"
    return ""


def extract_policy_term(text: str) -> str:
    m = re.search(r"Total\s+(\d+)\s+month\s+policy\s+premium", text, flags=re.IGNORECASE)
    if not m:
        m = re.search(r"(\d+)\s+month\s+policy", text, flags=re.IGNORECASE)
    if not m:
        m = re.search(r"Policy\s+Term\s*:?\s*(\d+)\s*months?", text, flags=re.IGNORECASE)
    if m:
        return f"{m.group(1)}个月"
    return ""


def _window(lines: List[str], idx: int, before: int = 3, after: int = 3) -> List[str]:
    s = max(0, idx - before)
    e = min(len(lines), idx + after + 1)
    return lines[s:e]


def extract_liability(text: str) -> Dict[str, Any]:
    res = {"selected": False, "bi_per_person": "", "bi_per_accident": "", "pd": ""}
    lines = text.splitlines()

    # Bodily Injury limits (Travelers sometimes shows limit above a line labeled just 'Liability')
    for i, line in enumerate(lines):
        if re.search(r"Bodily\s+Injury\s+Liability", line, flags=re.IGNORECASE) or \
           re.search(r"Liability\s+to\s+Others", line, flags=re.IGNORECASE) or \
           re.fullmatch(r"\s*Liability\s*", line, flags=re.IGNORECASE):
            for w in _window(lines, i, 3, 3):
                m = re.search(BI_PAIR_RE, w)
                if m:
                    res["bi_per_person"] = normalize_money(m.group(1))
                    res["bi_per_accident"] = normalize_money(m.group(2))
                    res["selected"] = True
                    break

    # Property Damage limit (may be plain integer like 25,000)
    for i, line in enumerate(lines):
        if re.search(r"Property\s+Damage\s*(Liability)?\b", line, flags=re.IGNORECASE):
            for w in _window(lines, i, 3, 3):
                m = re.search(MONEY_RE, w)
                if not m:
                    m = re.search(r"\b\d{2,5}(?:,\d{3})?\b(?!\.\d{2})", w)  # allow plain integer
                if m:
                    res["pd"] = normalize_money(m.group(0))
                    res["selected"] = True
                    break

    return res


def _find_nearby_amount(lines: List[str], idx: int, before: int = 3, after: int = 3) -> str:
    for w in _window(lines, idx, before, after):
        m = re.search(MONEY_RE, w)
        if m:
            return normalize_money(m.group(0))
    return ""


def extract_uninsured_motorist(text: str) -> Dict[str, Any]:
    lines = text.splitlines()
    umb = {"selected": False, "bi_per_person": "", "bi_per_accident": "", "pd": "", "deductible": "250"}

    # UMBI
    for i, line in enumerate(lines):
        if any(k.lower() in line.lower() for k in UMBI_KEYS):
            for w in _window(lines, i, 3, 5):
                m = re.search(BI_PAIR_RE, w)
                if m:
                    umb["bi_per_person"] = normalize_money(m.group(1))
                    umb["bi_per_accident"] = normalize_money(m.group(2))
                    umb["selected"] = True
                    break

    # UMPD — prefer plain integer amount near keyword to avoid premiums with cents
    for i, line in enumerate(lines):
        if any(k.lower() in line.lower() for k in UMPD_KEYS):
            for w in _window(lines, i, 3, 3):
                pm = re.search(r"\b\d{2,5}(?:,\d{3})?\b(?!\.\d{2})", w)
                if pm:
                    umb["pd"] = normalize_money(pm.group(0))
                    umb["selected"] = True
                    # Deductible near
                    dm = re.search(r"Deductible\s*\$?(\d{2,4})", "\n".join(_window(lines, i, 3, 3)), flags=re.IGNORECASE)
                    if dm:
                        umb["deductible"] = str(int(dm.group(1)))
                    break

    if not (umb["bi_per_person"] or umb["pd"]):
        umb["selected"] = False

    return umb


def extract_medical_payment(text: str) -> Dict[str, Any]:
    # Must have both keyword and an amount
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if re.search(r"Medical\s+Payments?", line, flags=re.IGNORECASE) or re.search(r"Med\s*Pay", line, flags=re.IGNORECASE):
            amt = _find_nearby_amount(lines, i, 3, 3)
            if amt:
                return {"selected": True, "med": amt}
            else:
                return {"selected": False, "med": ""}
    return {"selected": False, "med": ""}


def extract_personal_injury(text: str) -> Dict[str, Any]:
    # Must have both keyword and an amount
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if re.search(r"Personal\s+Injury\s+Protection|\bPIP\b", line, flags=re.IGNORECASE):
            amt = _find_nearby_amount(lines, i, 3, 3)
            if amt:
                return {"selected": True, "pip": amt}
            else:
                return {"selected": False, "pip": ""}
    return {"selected": False, "pip": ""}


def _looks_like_model(line: str) -> bool:
    s = line.strip()
    if not s:
        return False
    # Prefer lines like: 2018 TOYOTA RAV4 4 DOOR WAGON
    if re.match(rf"^{YEAR_RE}\s+[A-Z0-9][A-Z0-9\- ]+", s):
        if not any(sw in s.lower() for sw in ADDR_STOP_WORDS):
            return True
    # Also accept uppercase make/model without year when right above VIN
    if s.isupper() and len(s.split()) >= 2 and not any(sw in s.lower() for sw in ADDR_STOP_WORDS):
        return True
    # Accept short year+make on one line (e.g., '2018 BMW') which will be combined with next line
    if re.match(rf"^{YEAR_RE}\s+[A-Z]{2,}$", s):
        return True
    return False


def extract_vehicles(text: str) -> List[Dict[str, Any]]:
    vehicles: List[Dict[str, Any]] = []
    vin_pattern = re.compile(rf"(?:VIN[:\s#]*|)\b{VIN_RE}\b", re.IGNORECASE)
    lines = text.splitlines()

    for i, line in enumerate(lines):
        m = vin_pattern.search(line)
        if not m:
            continue
        vin = m.group(1)

        # model line: search up to 5 lines above, prefer year-leading line
        model = "未知车型"
        chosen_j = None
        for j in range(i - 1, max(-1, i - 6), -1):
            cand = lines[j].strip()
            if _looks_like_model(cand):
                model = re.sub(r"\s{2,}", " ", cand)
                chosen_j = j
                break
        # Combine with the next line if previous line is short like '2018 BMW' and next looks like model trim
        if chosen_j is not None and re.match(rf"^{YEAR_RE}\s+[A-Z]{2,}$", lines[chosen_j].strip()) and chosen_j + 1 < len(lines):
            nxt = lines[chosen_j + 1].strip()
            if nxt.isupper() and len(nxt) >= 3:
                model = f"{lines[chosen_j].strip()} {nxt}"

        # Also check same line left side
        if model == "未知车型":
            left = line.split(vin)[0].strip()
            if _looks_like_model(left):
                model = re.sub(r"\s{2,}", " ", left)

        # Build a local block around the VIN to find coverages
        block = "\n".join(lines[max(0, i - 20): min(len(lines), i + 21)])

        vehicle = {
            "model": model,
            "vin": vin,
            "collision": extract_deductible_bidirectional(block, "Collision"),
            "comprehensive": extract_deductible_bidirectional(block, "Comprehensive"),
            "rental": extract_limit_bidirectional(block, "Rental"),
            "roadside": extract_presence_bidirectional(block, "Roadside Assistance"),
        }
        vehicles.append(vehicle)

    # De-duplicate by VIN (keep first occurrence)
    seen = set()
    deduped: List[Dict[str, Any]] = []
    for v in vehicles:
        if v["vin"] in seen:
            continue
        seen.add(v["vin"])
        deduped.append(v)
    return deduped


def extract_deductible_bidirectional(text: str, keyword: str) -> Dict[str, Any]:
    result = {"selected": False, "deductible": ""}
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if keyword.lower() in line.lower():
            for w in _window(lines, i, 3, 3):
                # Prefer plain integer deductible near keyword (avoid picking premium $xx.xx)
                m_plain = re.search(r"\b\d{2,5}(?:,\d{3})?\b(?!\.\d{2})", w)
                if m_plain:
                    result["selected"] = True
                    result["deductible"] = re.sub(r"[^\d]", "", m_plain.group(0))
                    return result
                # Fallback: 'Deductible $500' style
                m = re.search(r"Deductible\s*:?\\s*(\\$?\\d{2,5}(?:,\\d{3})?)", w, flags=re.IGNORECASE)
                if m:
                    result["selected"] = True
                    result["deductible"] = re.sub(r"[^\d]", "", m.group(1))
                    return result
    return result


def extract_limit_bidirectional(text: str, keyword: str) -> Dict[str, Any]:
    result = {"selected": False, "limit": ""}
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if keyword.lower() in line.lower():
            for w in _window(lines, i, 3, 3):
                m = re.search(LIMIT_RE, w)
                if m:
                    result["selected"] = True
                    result["limit"] = re.sub(r"\s", "", m.group(0))
                    return result
    return result


def extract_presence_bidirectional(text: str, keyword: str) -> Dict[str, Any]:
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if keyword.lower() in line.lower():
            # presence with any amount or integer near it implies purchased
            for w in _window(lines, i, 3, 3):
                if re.search(MONEY_RE, w) or re.search(r"\b\d{1,4}\b", w):
                    return {"selected": True}
            return {"selected": True}  # some carriers show as a checkbox only
    # Also accept 'Roadside Assistance Coverage' variant
    if "roadside assistance coverage" in text.lower():
        return {"selected": True}
    return {"selected": False}
