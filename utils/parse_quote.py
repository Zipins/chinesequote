import re
import io
import fitz  # PyMuPDF
import boto3
from PIL import Image

def extract_text_from_pdf_or_image(file):
    if file.name.lower().endswith(".pdf"):
        with fitz.open(stream=file.read(), filetype="pdf") as doc:
            text = ""
            for page in doc:
                text += page.get_text()
            if not text.strip():
                # 若为扫描件，尝试用 OCR
                file.seek(0)
                images = []
                with fitz.open(stream=file.read(), filetype="pdf") as pdf:
                    for page in pdf:
                        pix = page.get_pixmap(dpi=300)
                        img_data = pix.tobytes("png")
                        images.append(img_data)
                return run_textract_ocr(images)
            return text
    elif file.name.lower().endswith((".png", ".jpg", ".jpeg")):
        return run_textract_ocr([file.read()])
    else:
        raise ValueError("不支持的文件类型")

def run_textract_ocr(image_bytes_list):
    textract = boto3.client("textract")
    full_text = ""
    for img_bytes in image_bytes_list:
        response = textract.detect_document_text(Document={"Bytes": img_bytes})
        blocks = response.get("Blocks", [])
        page_text = "\n".join([b["Text"] for b in blocks if b["BlockType"] == "LINE"])
        full_text += page_text + "\n"
    return full_text.strip()

def extract_company_name(text):
    text_lower = text.lower()
    if "progressive" in text_lower:
        return "Progressive"
    elif "travelers" in text_lower:
        return "Travelers"
    elif "state farm" in text_lower:
        return "State Farm"
    elif "geico" in text_lower:
        return "GEICO"
    elif "allstate" in text_lower:
        return "Allstate"
    return "某保险公司"

def extract_quote_data(file, return_raw_text=False):
    full_text = extract_text_from_pdf_or_image(file)
    lines = full_text.splitlines()

    data = {
        "company": extract_company_name(full_text),
        "total_premium": "",
        "policy_term": "6个月",
        "liability": {"selected": False, "bi_per_person": "", "bi_per_accident": "", "pd": ""},
        "uninsured_motorist": {"selected": False, "bi_per_person": "", "bi_per_accident": "", "pd": "", "deductible": "250"},
        "medical_payment": {"selected": False, "med": ""},
        "personal_injury": {"selected": False, "pip": ""},
        "vehicles": [],
    }

    # 总保费
    match = re.search(r"Total\s+\d+\s*month.*?\$[\d,]+\.\d{2}", full_text, re.IGNORECASE)
    if match:
        price = re.search(r"\$[\d,]+\.\d{2}", match.group())
        if price:
            data["total_premium"] = price.group()
        term = re.search(r"Total\s+(\d+)\s*month", match.group(), re.IGNORECASE)
        if term:
            data["policy_term"] = f"{term.group(1)}个月"

    # 责任险 Liability to Others
    liab = re.search(r"Bodily Injury Liability\s*\$([\d,]+)/\$([\d,]+).*?Property Damage Liability\s*\$([\d,]+)", full_text, re.IGNORECASE)
    if liab:
        data["liability"]["selected"] = True
        data["liability"]["bi_per_person"] = f"${liab.group(1)}"
        data["liability"]["bi_per_accident"] = f"${liab.group(2)}"
        data["liability"]["pd"] = f"${liab.group(3)}"

    # Uninsured Motorist
    umbi = re.search(r"Uninsured.*?Bodily Injury\s*\$([\d,]+)/\$([\d,]+)", full_text, re.IGNORECASE)
    umpd = re.search(r"Uninsured.*?Property Damage\s*\$([\d,]+)", full_text, re.IGNORECASE)
    if umbi or umpd:
        data["uninsured_motorist"]["selected"] = True
        if umbi:
            data["uninsured_motorist"]["bi_per_person"] = f"${umbi.group(1)}"
            data["uninsured_motorist"]["bi_per_accident"] = f"${umbi.group(2)}"
        if umpd:
            data["uninsured_motorist"]["pd"] = f"${umpd.group(1)}"

    # Medical Payment
    med = re.search(r"Medical Payments.*?\$([\d,]+)", full_text, re.IGNORECASE)
    if med:
        data["medical_payment"]["selected"] = True
        data["medical_payment"]["med"] = f"${med.group(1)}"

    # Personal Injury
    pip = re.search(r"Personal Injury Protection.*?\$([\d,]+)", full_text, re.IGNORECASE)
    if pip:
        data["personal_injury"]["selected"] = True
        data["personal_injury"]["pip"] = f"${pip.group(1)}"

    # 车辆信息
    for i, line in enumerate(lines):
        if re.search(r"\bVIN[:：]?\s*([A-HJ-NPR-Z0-9]{17})", line):
            vin_match = re.search(r"\bVIN[:：]?\s*([A-HJ-NPR-Z0-9]{17})", line)
            vin = vin_match.group(1)
            model = lines[i - 1].strip() if i > 0 else "未知车型"
            vehicle = {
                "model": model,
                "vin": vin,
                "collision": {"selected": False, "deductible": ""},
                "comprehensive": {"selected": False, "deductible": ""},
                "rental": {"selected": False, "limit": ""},
                "roadside": {"selected": False},
            }

            # 搜索后几行保障信息
            for j in range(i + 1, min(i + 15, len(lines))):
                l = lines[j]
                if "Collision" in l and "$" in l:
                    vehicle["collision"]["selected"] = True
                    deduct = re.search(r"\$([\d,]+)", l)
                    if deduct:
                        vehicle["collision"]["deductible"] = deduct.group(1)
                elif "Comprehensive" in l and "$" in l:
                    vehicle["comprehensive"]["selected"] = True
                    deduct = re.search(r"\$([\d,]+)", l)
                    if deduct:
                        vehicle["comprehensive"]["deductible"] = deduct.group(1)
                elif "Rental" in l and re.search(r"\d+/\d+", l):
                    vehicle["rental"]["selected"] = True
                    lim = re.search(r"\d+/\d+", l)
                    if lim:
                        vehicle["rental"]["limit"] = lim.group()
                elif "Roadside" in l:
                    vehicle["roadside"]["selected"] = True
                elif "VIN" in l:
                    break  # 下一辆车开始

            data["vehicles"].append(vehicle)

    return (data, full_text) if return_raw_text else data
