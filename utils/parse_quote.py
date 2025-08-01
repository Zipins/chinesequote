import boto3
import re
import io
import fitz  # PyMuPDF
from PIL import Image
import traceback
from copy import deepcopy

def extract_quote_data(file, return_raw_text=False):
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
            "vehicles": extract_vehicles(full_text)
        }

        return (data, full_text) if return_raw_text else data

    except Exception:
        traceback.print_exc()
        raise

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

def extract_company_name(text):
    known_companies = ["Progressive", "Travelers", "Allstate", "Geico", "Liberty Mutual", "State Farm", "Safeco", "Nationwide"]
    for name in known_companies:
        if name.lower() in text.lower():
            return name
    return "某保险公司"

def extract_total_premium(text):
    match = re.search(r"pay.*?premium.*?\$([\d,]+\.\d{2})", text, re.IGNORECASE)
    if match:
        return f"${match.group(1)}"
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if "pay" in line.lower() and "premium" in line.lower():
            for j in range(i, min(i + 5, len(lines))):
                m = re.search(r"\$([\d,]+\.\d{2})", lines[j])
                if m:
                    return f"${m.group(1)}"
    return ""

def extract_policy_term(text):
    match = re.search(r"Total\s+(\d+)\s+month", text, re.IGNORECASE)
    if not match:
        match = re.search(r"(\d+)\s+month\s+policy", text, re.IGNORECASE)
    if match:
        return f"{match.group(1)}个月"
    return ""

def extract_liability(text):
    result = {"selected": False, "bi_per_person": "", "bi_per_accident": "", "pd": ""}
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if "liability" in line.lower():
            for j in range(i + 1, min(i + 5, len(lines))):
                match = re.search(r"(\d{1,3}[,\d]*)/(\d{1,3}[,\d]*)", lines[j])
                if match:
                    result["bi_per_person"] = f"${match.group(1)}"
                    result["bi_per_accident"] = f"${match.group(2)}"
                    result["selected"] = True
                    break
    for i, line in enumerate(lines):
        if "property damage" in line.lower():
            for j in range(i, min(i+3, len(lines))):
                pd_match = re.search(r"\$?(\d{1,3}[,\d]*)", lines[j])
                if pd_match:
                    result["pd"] = f"${pd_match.group(1)}"
                    result["selected"] = True
                    break
    return result

def extract_uninsured_motorist(text):
    result = {"selected": False, "bi_per_person": "", "bi_per_accident": "", "pd": "", "deductible": "250"}
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if "unins" in line.lower() and "/" in line:
            match = re.search(r"(\d{1,3}[,\d]*)/(\d{1,3}[,\d]*)", line)
            if match:
                result["bi_per_person"] = f"${match.group(1)}"
                result["bi_per_accident"] = f"${match.group(2)}"
                result["selected"] = True
        if "unins" in line.lower() and "pd" in line.lower():
            pd_match = re.search(r"\$?(\d{1,3}[,\d]*)", line)
            if pd_match:
                result["pd"] = f"${pd_match.group(1)}"
                result["selected"] = True
    return result

def extract_medical_payment(text):
    result = {"selected": False, "med": ""}
    match = re.search(r"Medical Payments\s*\$?([\d,]+)", text)
    if match:
        result["selected"] = True
        result["med"] = f"${match.group(1)}"
    return result

def extract_personal_injury(text):
    result = {"selected": False, "pip": ""}
    match = re.search(r"Personal Injury Protection\s*\$?([\d,]+)", text)
    if match:
        result["selected"] = True
        result["pip"] = f"${match.group(1)}"
    return result

def extract_vehicles(text):
    vehicles = []
    vin_pattern = re.compile(r"(VIN[:\s]*)?([A-HJ-NPR-Z0-9]{17})", re.IGNORECASE)
    lines = text.splitlines()
    for i, line in enumerate(lines):
        vin_match = vin_pattern.search(line)
        if vin_match:
            vin = vin_match.group(2)
            model = "未知车型"
            for j in range(i-1, max(i-5, -1), -1):
                model_line = lines[j].strip()
                if re.match(r"\d{4}\s+[A-Z0-9 ]{3,}", model_line):
                    model = model_line
                    break
            block_text = "\n".join(lines[max(i-5, 0):i+15])
            vehicle = {
                "model": model.strip(),
                "vin": vin.strip(),
                "collision": extract_deductible_multi(block_text, "Collision"),
                "comprehensive": extract_deductible_multi(block_text, "Comprehensive"),
                "rental": extract_presence_multi(block_text, "Rental"),
                "roadside": extract_presence_multi(block_text, "Roadside Assistance")
            }
            vehicles.append(vehicle)
    return vehicles

def extract_deductible_multi(text, keyword):
    result = {"selected": False, "deductible": ""}
    lines = text.splitlines()
    for line in lines:
        if keyword.lower() in line.lower():
            match = re.search(rf"{keyword}[^\d]*?(\d{{1,4}})", line, re.IGNORECASE)
            if match:
                result["selected"] = True
                result["deductible"] = match.group(1)
                break
    return result

def extract_presence_multi(text, keyword):
    lines = text.splitlines()
    for line in lines:
        if keyword.lower() in line.lower():
            if re.search(r"\$?\d{1,4}(\.\d{2})?", line):
                return {"selected": True}
    return {"selected": False}
