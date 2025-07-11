import boto3
import re
import io
import fitz  # PyMuPDF
from PIL import Image


def extract_quote_data(file, return_raw_text=False):
    file_bytes = file.read()
    file.seek(0)
    file_suffix = file.name.split(".")[-1].lower()

    textract = boto3.client("textract", region_name="us-east-1")

    # 判断 PDF 是文本型还是扫描型
    def is_pdf_text_based(pdf_bytes):
        try:
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            for page in doc:
                if page.get_text().strip():
                    return True
            return False
        except Exception:
            return False

    if file_suffix == "pdf":
        if is_pdf_text_based(file_bytes):
            # 文本型 PDF → 直接上传给 Textract
            response = textract.detect_document_text(Document={"Bytes": file_bytes})
        else:
            # 扫描件 PDF → 转为图片后上传
            images = pdf_to_images(file_bytes)
            if not images:
                raise ValueError("PDF 转图片失败")
            response = textract.detect_document_text(Document={"Bytes": images[0]})

    elif file_suffix in ["jpg", "jpeg", "png"]:
        response = textract.detect_document_text(Document={"Bytes": file_bytes})
    else:
        raise ValueError("不支持的文件格式")

    lines = [block["Text"] for block in response["Blocks"] if block["BlockType"] == "LINE"]
    full_text = "\n".join(lines)

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


def pdf_to_images(pdf_bytes):
    images = []
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    for page in doc:
        pix = page.get_pixmap(dpi=300)
        img_bytes = pix.tobytes("png")
        images.append(img_bytes)
    return images

# 以下为原有字段提取函数不变

def extract_company_name(text):
    if "Progressive" in text:
        return "Progressive"
    if "Travelers" in text:
        return "Travelers"
    return "某保险公司"

def extract_total_premium(text):
    match = re.search(r"Total\s+\d+\s+month.*?\$([\d,]+\.\d{2})", text, re.IGNORECASE)
    if match:
        return f"${match.group(1)}"
    return ""

def extract_policy_term(text):
    match = re.search(r"Total\s+(\d+)\s+month", text, re.IGNORECASE)
    if match:
        return f"{match.group(1)}个月"
    return ""

def extract_liability(text):
    result = {
        "selected": False,
        "bi_per_person": "",
        "bi_per_accident": "",
        "pd": ""
    }
    bi_match = re.search(r"Bodily Injury Liability\s*\$([\d,]+)[^\d]+([\d,]+)", text)
    pd_match = re.search(r"Property Damage Liability\s*\$([\d,]+)", text)
    if bi_match and pd_match:
        result["selected"] = True
        result["bi_per_person"] = f"${bi_match.group(1)}"
        result["bi_per_accident"] = f"${bi_match.group(2)}"
        result["pd"] = f"${pd_match.group(1)}"
    return result

def extract_uninsured_motorist(text):
    result = {
        "selected": False,
        "bi_per_person": "",
        "bi_per_accident": "",
        "pd": "",
        "deductible": "250"
    }
    bi_match = re.search(r"Uninsured/Underinsured Motorist Bodily Injury\s*\$([\d,]+)[^\d]+([\d,]+)", text)
    pd_match = re.search(r"Uninsured/Underinsured Motorist Property Damage\s*\$([\d,]+)", text)
    if bi_match or pd_match:
        result["selected"] = True
        if bi_match:
            result["bi_per_person"] = f"${bi_match.group(1)}"
            result["bi_per_accident"] = f"${bi_match.group(2)}"
        if pd_match:
            result["pd"] = f"${pd_match.group(1)}"
    return result

def extract_medical_payment(text):
    result = {"selected": False, "med": ""}
    match = re.search(r"Medical Payments\s*\$([\d,]+)", text)
    if match:
        result["selected"] = True
        result["med"] = f"${match.group(1)}"
    return result

def extract_personal_injury(text):
    result = {"selected": False, "pip": ""}
    match = re.search(r"Personal Injury Protection\s*\$([\d,]+)", text)
    if match:
        result["selected"] = True
        result["pip"] = f"${match.group(1)}"
    return result

def extract_vehicles(text):
    vehicles = []
    vehicle_blocks = re.split(r"(?=VIN[:\s])", text)
    for block in vehicle_blocks:
        vin_match = re.search(r"VIN[:\s]*([A-HJ-NPR-Z0-9]{17})", block)
        if not vin_match:
            continue
        vin = vin_match.group(1)
        model_line = extract_model_line(block, vin)
        vehicle = {
            "model": model_line.strip(),
            "vin": vin,
            "collision": extract_deductible(block, "Collision"),
            "comprehensive": extract_deductible(block, "Comprehensive"),
            "rental": extract_rental(block),
            "roadside": extract_presence(block, "Roadside Assistance"),
        }
        vehicles.append(vehicle)
    return vehicles

def extract_model_line(block, vin):
    lines = block.strip().split("\n")
    for i, line in enumerate(lines):
        if vin in line and i > 0:
            return lines[i - 1]
    return "未知车型"

def extract_deductible(text, keyword):
    result = {"selected": False, "deductible": ""}
    match = re.search(fr"{keyword}.*?\$([\d,]+)", text, re.IGNORECASE)
    if match:
        result["selected"] = True
        result["deductible"] = match.group(1)
    return result

def extract_rental(text):
    result = {"selected": False, "limit": ""}
    match = re.search(r"Rental.*?([\d,]+)/([\d,]+)", text)
    if match:
        result["selected"] = True
        result["limit"] = f"{match.group(1)}/{match.group(2)}"
    return result

def extract_presence(text, keyword):
    return {"selected": keyword.lower() in text.lower()}
