# utils/parse_quote.py
import boto3
import os
import io
import re
import fitz  # PyMuPDF
from PIL import Image
from typing import Dict

def extract_quote_data(uploaded_file) -> Dict:
    file_bytes = uploaded_file.read()
    filename = uploaded_file.name
    text = extract_text_from_textract(file_bytes, filename)

    data = {
        "company": extract_company_name(text),
        "total_premium": extract_premium(text),
        "policy_term": extract_term(text),
        "liability": extract_liability(text),
        "uninsured_motorist": extract_uninsured(text),
        "medical_payment": extract_medpay(text),
        "personal_injury": extract_pip(text),
        "vehicles": extract_vehicles(text),
    }
    return data

def extract_text_from_textract(file_bytes: bytes, filename: str) -> str:
    ext = filename.lower().split(".")[-1]
    client = boto3.client("textract", region_name=os.getenv("AWS_REGION", "us-east-1"))

    # PDF：每页渲染成图片上传（解决 UnsupportedDocumentException）
    if ext == "pdf":
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        text = ""
        for page in doc:
            pix = page.get_pixmap(dpi=300)
            image_bytes = pix.tobytes("png")
            response = client.detect_document_text(Document={'Bytes': image_bytes})
            blocks = response.get("Blocks", [])
            page_text = "\n".join(block["Text"] for block in blocks if block["BlockType"] == "LINE")
            text += page_text + "\n"
        return text

    # PNG/JPG 图像处理
    elif ext in ["png", "jpg", "jpeg"]:
        image = Image.open(io.BytesIO(file_bytes)).convert("RGB")
        buffer = io.BytesIO()
        image.save(buffer, format="PNG")
        textract_doc = {'Bytes': buffer.getvalue()}
        response = client.detect_document_text(Document=textract_doc)
        blocks = response.get("Blocks", [])
        text = "\n".join(block["Text"] for block in blocks if block["BlockType"] == "LINE")
        return text

    else:
        raise ValueError("不支持的文件类型")

# ✅ 智能保险公司识别
def extract_company_name(text: str) -> str:
    known_companies = [
        "Progressive", "Travelers", "Safeco", "Allstate", "State Farm",
        "GEICO", "Liberty Mutual", "Nationwide", "Bristol West",
        "Mercury", "Amica", "Hartford", "Kemper", "Infinity"
    ]
    for company in known_companies:
        if company.lower() in text.lower():
            return company

    match = re.search(r"(?:Underwritten by|Quote from|Provided by)[\s:]*([A-Za-z0-9 &.,\-]+)", text, re.IGNORECASE)
    if match:
        raw_name = match.group(1).strip()
        clean_name = re.sub(r"\b(Inc|Co|LLC|Ltd)\b\.?", "", raw_name).strip()
        return clean_name

    return "某保险公司"

def extract_term(text):
    match = re.search(r"Total\s+(\d+)\s+month", text, re.IGNORECASE)
    return f"{match.group(1)}个月" if match else "6个月"

def extract_premium(text):
    match = re.search(r"Total\s+\d+\s+month[^$]*\$\s*([\d,]+\.\d+)", text, re.IGNORECASE)
    return f"${match.group(1)}" if match else "未识别"

def extract_liability(text):
    result = {"selected": False, "bi_per_person": "", "bi_per_accident": "", "pd": ""}
    if "Liability to Others" in text or "Bodily Injury Liability" in text:
        bi = re.search(r"Bodily Injury Liability.*?\$([\d,]+)[^\d]+[\$/]*([\d,]+)", text)
        pd = re.search(r"Property Damage Liability.*?\$([\d,]+)", text)
        if bi and pd:
            result["selected"] = True
            result["bi_per_person"] = bi.group(1)
            result["bi_per_accident"] = bi.group(2)
            result["pd"] = pd.group(1)
    return result

def extract_uninsured(text):
    result = {"selected": False, "bi_per_person": "", "bi_per_accident": "", "pd": "", "deductible": "250"}
    bi = re.search(r"Uninsured.*?Bodily Injury.*?\$([\d,]+)[^\d]+[\$/]*([\d,]+)", text, re.IGNORECASE)
    pd = re.search(r"Uninsured.*?Property Damage.*?\$([\d,]+)", text, re.IGNORECASE)
    if bi:
        result["selected"] = True
        result["bi_per_person"] = bi.group(1)
        result["bi_per_accident"] = bi.group(2)
    if pd:
        result["pd"] = pd.group(1)
    return result

def extract_medpay(text):
    result = {"selected": False, "med": ""}
    match = re.search(r"Medical Payments.*?\$([\d,]+)", text, re.IGNORECASE)
    if match:
        result["selected"] = True
        result["med"] = match.group(1)
    return result

def extract_pip(text):
    result = {"selected": False, "pip": ""}
    match = re.search(r"Personal Injury Protection.*?\$([\d,]+)", text, re.IGNORECASE)
    if match:
        result["selected"] = True
        result["pip"] = match.group(1)
    return result

def extract_vehicles(text):
    vehicles = []
    vehicle_blocks = re.findall(r"(\d{4}\s+[A-Z0-9\- ]+)\s+VIN[:：]?\s*([A-HJ-NPR-Z0-9]{10,})", text)
    for model_str, vin in vehicle_blocks:
        vehicles.append({
            "model": model_str.strip(),
            "vin": vin.strip(),
            "collision": {"selected": True, "deductible": "495"},
            "comprehensive": {"selected": True, "deductible": "495"},
            "rental": {"selected": True, "limit": "40/30"},
            "roadside": {"selected": True},
        })
    return vehicles
