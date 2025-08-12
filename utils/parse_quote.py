
import re
import io
import traceback
from copy import deepcopy
from typing import Dict, Any, List

import fitz  # PyMuPDF
from PIL import Image

try:
    import boto3
    from botocore.exceptions import BotoCoreError, ClientError
except Exception:
    boto3 = None
    BotoCoreError = Exception  # type: ignore
    ClientError = Exception  # type: ignore

YEAR_RE = r"(19\d{2}|20\d{2})"
VIN_RE = r"([A-HJ-NPR-Z\d]{17})"
MONEY_RE = r"\$?\d{1,3}(?:,\d{3})*(?:\.\d{2})?"
BI_PAIR_RE = r"(\$?\d{1,3}(?:,\d{3})*)(?:\s*/\s*)(\$?\d{1,3}(?:,\d{3})*)"
LIMIT_RE = r"\d{1,3}(?:,\d{3})?\s*/\s*\d{1,4}(?:,\d{3})?"

UMBI_KEYS = [
    "Uninsured/Underinsured Motorist Bodily Injury",
    "Uninsured Motorist Bodily Injury",
    "Underinsured Motorist Bodily Injury",
    "UMBI",
    "Uninsd/Underinsd Motorists",
    "Uninsd Motorists",
    "Underinsd Motorists",
]
UMPD_KEYS = [
    "Uninsured/Underinsured Motorist Property Damage",
    "Uninsured Motorist Property Damage",
    "Underinsured Motorist Property Damage",
    "UMPD",
    "Uninsd/Underinsd Motorists PD",
    "Uninsd Motorists PD",
    "Underinsd Motorists PD",
]

ADDR_STOP_WORDS = [
    "street","st.","st ","road","rd.","rd ","ave","avenue","boulevard","blvd","lane","ln",
    "drive","dr","suite","ste","apt","unit"
]

def _get_textract_client(region: str = "us-east-1"):
    if boto3 is None:
        return None
    try:
        return boto3.client("textract", region_name=region)
    except Exception:
        return None

def _textract_analyze_tables(img_png_bytes: bytes, client) -> List[List[List[str]]]:
    if client is None:
        return []
    try:
        resp = client.analyze_document(
            Document={"Bytes": img_png_bytes},
            FeatureTypes=["TABLES", "FORMS"],
        )
    except Exception:
        return []
    blocks = resp.get("Blocks", [])
    id2block = {b["Id"]: b for b in blocks}
    tables = []
    for b in blocks:
        if b.get("BlockType") != "TABLE":
            continue
        rows_map: Dict[int, Dict[int, str]] = {}
        for rel in b.get("Relationships", []):
            if rel.get("Type") != "CHILD":
                continue
            for cid in rel.get("Ids", []):
                cell = id2block.get(cid)
                if not cell or cell.get("BlockType") != "CELL":
                    continue
                r = cell.get("RowIndex", 0)
                c = cell.get("ColumnIndex", 0)
                text_parts: List[str] = []
                for cr in cell.get("Relationships", []):
                    if cr.get("Type") != "CHILD":
                        continue
                    for wid in cr.get("Ids", []):
                        wb = id2block.get(wid)
                        if not wb: 
                            continue
                        if wb.get("BlockType") == "WORD":
                            text_parts.append(wb.get("Text",""))
                        elif wb.get("BlockType") == "SELECTION_ELEMENT" and wb.get("SelectionStatus") == "SELECTED":
                            text_parts.append("☑")
                rows_map.setdefault(r, {})[c] = " ".join(text_parts).strip()
        if rows_map:
            maxc = max((max(cols.keys()) for cols in rows_map.values()), default=0)
            norm: List[List[str]] = []
            for r in sorted(rows_map.keys()):
                norm.append([rows_map[r].get(c, "") for c in range(1, maxc+1)])
            tables.append(norm)
    return tables

def _textract_detect_lines(img_png_bytes: bytes, client) -> str:
    if client is None:
        return ""
    try:
        resp = client.detect_document_text(Document={"Bytes": img_png_bytes})
        lines = [b["Text"] for b in resp.get("Blocks", []) if b.get("BlockType") == "LINE"]
        return "\n".join(lines)
    except Exception:
        return ""

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

def detect_company(text: str) -> str:
    lowered = text.lower()
    if "progressive" in lowered: return "Progressive"
    if "travelers" in lowered: return "Travelers"
    if "allstate" in lowered: return "Allstate"
    if "geico" in lowered: return "Geico"
    if "liberty mutual" in lowered: return "Liberty Mutual"
    if "safeco" in lowered: return "Safeco"
    if "state farm" in lowered: return "State Farm"
    if "nationwide" in lowered: return "Nationwide"
    return "某保险公司"

def extract_company_name(text: str) -> str:
    return detect_company(text)

def pick_pif_monthly_down(text: str) -> Dict[str, str]:
    amts = [float(x.replace(",", "")) for x in re.findall(r"\$([\d,]+\.\d{2})", text)]
    out: Dict[str, str] = {}
    if len(amts) < 3:
        return out
    for i in range(len(amts)-2):
        a = sorted(amts[i:i+3])
        down, pif, monthly = a[0], a[1], a[2]
        if down < 600 and pif < monthly:
            out["down"] = f"${down:,.2f}"
            out["pay_in_full"] = f"${pif:,.2f}"
            out["monthly_total"] = f"${monthly:,.2f}"
            break
    return out

def extract_total_premium(text: str) -> str:
    patterns = [
        r"estimated\s+pay-?in-?full[\s\S]*?\$([\d,]+\.\d{2})",
        r"pay-?in-?full[\s\S]*?\$([\d,]+\.\d{2})",
        r"Total\s+\d+\s+month\s+policy\s+premium[\s\S]*?\$([\d,]+\.\d{2})",
        r"your\s+estimated\s+total\s+premium[\s\S]*?\$([\d,]+\.\d{2})",
        r"Total\s+policy\s+premium[\s\S]*?\$([\d,]+\.\d{2})",
    ]
    for p in patterns:
        m = re.search(p, text, flags=re.I)
        if m:
            span_start = m.start()
            guard = text[max(0, span_start-40): span_start+40].lower()
            if "savings" in guard:
                continue
            return normalize_money(m.group(1))
    pick = pick_pif_monthly_down(text)
    if pick.get("pay_in_full"):
        return pick["pay_in_full"]
    amounts = [(m.group(0), m.start()) for m in re.finditer(MONEY_RE, text)]
    cands = []
    for raw, pos in amounts:
        try:
            if "savings" in text[max(0, pos-30):pos+30].lower():
                continue
            val = float(re.sub(r"[,$]", "", raw))
            if val >= 1000:
                cands.append(val)
        except Exception:
            pass
    if cands:
        return f"${max(cands):,.2f}"
    return ""

def extract_policy_term(text: str) -> str:
    m = re.search(r"Total\s+(\d+)\s+month\s+policy\s+premium", text, flags=re.I)
    if not m:
        m = re.search(r"(\d+)\s+month\s+policy", text, flags=re.I)
    if not m:
        m = re.search(r"Policy\s+Term\s*:?\s*(\d+)\s*months?", text, flags=re.I)
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
    for i, line in enumerate(lines):
        if re.search(r"Bodily\s+Injury\s+Liability", line, flags=re.I) or \
           re.search(r"Liability\s+to\s+Others", line, flags=re.I) or \
           re.fullmatch(r"\s*Liability\s*", line, flags=re.I):
            for w in _window(lines, i, 3, 3):
                m = re.search(BI_PAIR_RE, w)
                if m:
                    res["bi_per_person"] = normalize_money(m.group(1))
                    res["bi_per_accident"] = normalize_money(m.group(2))
                    res["selected"] = True
                    break
    for i, line in enumerate(lines):
        if re.search(r"Property\s+Damage\s*(Liability)?\b", line, flags=re.I):
            for w in _window(lines, i, 3, 3):
                m = re.search(MONEY_RE, w) or re.search(r"\b\d{2,5}(?:,\d{3})?\b(?!\.\d{2})", w)
                if m:
                    res["pd"] = normalize_money(m.group(0))
                    res["selected"] = True
                    break
    return res

def _find_nearby_amount(lines: List[str], idx: int, before: int = 3, after: int = 3) -> str:
    for w in _window(lines, idx, before, after):
        m = re.search(MONEY_RE, w)
        if m: return normalize_money(m.group(0))
    return ""

def extract_uninsured_motorist(text: str) -> Dict[str, Any]:
    lines = text.splitlines()
    umb = {"selected": False, "bi_per_person": "", "bi_per_accident": "", "pd": "", "deductible": "250"}
    for i, line in enumerate(lines):
        if any(k.lower() in line.lower() for k in UMBI_KEYS):
            for w in _window(lines, i, 3, 5):
                m = re.search(BI_PAIR_RE, w)
                if m:
                    umb["bi_per_person"] = normalize_money(m.group(1))
                    umb["bi_per_accident"] = normalize_money(m.group(2))
                    umb["selected"] = True
                    break
    for i, line in enumerate(lines):
        if any(k.lower() in line.lower() for k in UMPD_KEYS):
            for w in _window(lines, i, 3, 3):
                pm = re.search(r"\b\d{2,5}(?:,\d{3})?\b(?!\.\d{2})", w)
                if pm:
                    umb["pd"] = normalize_money(pm.group(0))
                    umb["selected"] = True
                    break
    if not (umb["bi_per_person"] or umb["pd"]):
        umb["selected"] = False
    return umb

def extract_medical_payment(text: str) -> Dict[str, Any]:
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if re.search(r"Medical\s+Payments?", line, flags=re.I) or re.search(r"Med\s*Pay", line, flags=re.I):
            amt = _find_nearby_amount(lines, i, 3, 3)
            if amt:
                return {"selected": True, "med": amt}
            else:
                return {"selected": False, "med": ""}
    return {"selected": False, "med": ""}

def extract_personal_injury(text: str) -> Dict[str, Any]:
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if re.search(r"Personal\s+Injury\s+Protection|\bPIP\b", line, flags=re.I):
            amt = _find_nearby_amount(lines, i, 3, 3)
            if amt:
                return {"selected": True, "pip": amt}
            else:
                return {"selected": False, "pip": ""}
    return {"selected": False, "pip": ""}

def _looks_like_model(line: str) -> bool:
    s = line.strip()
    if not s: return False
    if re.match(rf"^{YEAR_RE}\s+[A-Z0-9][A-Z0-9\- ]+", s):
        if not any(sw in s.lower() for sw in ADDR_STOP_WORDS):
            return True
    if s.isupper() and len(s.split()) >= 2 and not any(sw in s.lower() for sw in ADDR_STOP_WORDS):
        return True
    if re.match(rf"^{YEAR_RE}\s+[A-Z]{2,}$", s):
        return True
    return False

def extract_vehicles(text: str) -> List[Dict[str, Any]]:
    vehicles: List[Dict[str, Any]] = []
    vin_pattern = re.compile(rf"(?:VIN[:\s#]*|)\b{VIN_RE}\b", re.I)
    lines = text.splitlines()
    for i, line in enumerate(lines):
        m = vin_pattern.search(line)
        if not m: continue
        vin = m.group(1)
        model = "未知车型"; hit = None
        for j in range(i-1, max(-1, i-6), -1):
            cand = lines[j].strip()
            if _looks_like_model(cand):
                model, hit = re.sub(r"\s{2,}", " ", cand), j
                break
        if hit is not None and re.match(rf"^{YEAR_RE}\s+[A-Z]{2,}$", lines[hit].strip()) and hit+1 < len(lines):
            nxt = lines[hit+1].strip()
            if nxt.isupper() and len(nxt) >= 3:
                model = f"{lines[hit].strip()} {nxt}"
        if model == "未知车型":
            left = line.split(vin)[0].strip()
            if _looks_like_model(left):
                model = re.sub(r"\s{2,}", " ", left)
        block = "\n".join(lines[max(0, i-20): min(len(lines), i+21)])
        vehicles.append({
            "model": model,
            "vin": vin,
            "collision": extract_deductible_bidirectional(block, "Collision"),
            "comprehensive": extract_deductible_bidirectional(block, "Comprehensive"),
            "rental": extract_limit_bidirectional(block, "Rental"),
            "roadside": extract_presence_bidirectional(block, "Roadside Assistance"),
        })
    seen, out = set(), []
    for v in vehicles:
        if v["vin"] in seen: continue
        seen.add(v["vin"]); out.append(v)
    return out

def extract_deductible_bidirectional(text: str, keyword: str) -> Dict[str, Any]:
    result = {"selected": False, "deductible": ""}
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if keyword.lower() in line.lower():
            for w in _window(lines, i, 3, 3):
                m_plain = re.search(r"\b\d{2,5}(?:,\d{3})?\b(?!\.\d{2})", w)
                if m_plain:
                    result["selected"] = True
                    result["deductible"] = re.sub(r"[^\d]", "", m_plain.group(0))
                    return result
                m = re.search(r"Deductible\s*:?\\s*(\\$?\\d{2,5}(?:,\\d{3})?)", w, flags=re.I)
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
            for w in _window(lines, i, 3, 3):
                if re.search(MONEY_RE, w) or re.search(r"\b\d{1,4}\b", w):
                    return {"selected": True}
            return {"selected": True}
    if "roadside assistance coverage" in text.lower():
        return {"selected": True}
    return {"selected": False}

