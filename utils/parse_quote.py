# utils/parse_quote.py
import boto3
import os
import io
import re
import fitz  # PyMuPDF
from PIL import Image
from typing import Dict
from botocore.exceptions import BotoCoreError, ClientError

def is_scanned_pdf(file_bytes: bytes) -> bool:
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    for page in doc:
        if page.get_text().strip():
            return False  # 有文本
    return True  # 全部是图片

def extract_text_from_textract(file_bytes: bytes, filename: str) -> str:
    ext = filename.lower().split(".")[-1]
    client = boto3.client("textract", region_name=os.getenv("AWS_REGION", "us-east-1"))

    if ext == "pdf":
        if is_scanned_pdf(file_bytes):
            # 调用 Textract StartDocumentTextDetection（异步）
            response = client.analyze_document(
                Document={'Bytes': file_bytes},
                FeatureTypes=["TABLES", "FORMS"]
            )
        else:
            # 文本型 PDF 用 detect_document_text
            response = client.detect_document_text(Document={'Bytes': file_bytes})
    else:
        image = Image.open(io.BytesIO(file_bytes)).convert("RGB")
        buffer = io.BytesIO()
        image.save(buffer, format="PNG")
        response = client.detect_document_text(Document={'Bytes': buffer.getvalue()})

    blocks = response.get("Blocks", [])
    text = "\n".join(block["Text"] for block in blocks if block["BlockType"] == "LINE")
    return text

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

# 以下为解析逻辑，支持 Progressive 和 Travelers（可扩展）
def extract_company_name(text):
    if "Progressive" in text:
        return "Progressive"
    if "Travelers" in text:
        return "Travelers"
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
    vin_blocks = re.findall(r"(VIN[:：]?\s*[A-HJ-NPR-Z0-9]{10,})", text)
    for vin in vin_blocks:
        vin_clean = vin.split(":")[-1].strip()
        model_match = re.search(r"(\d{4})\s+([A-Z][A-Z0-9\- ]+)", text)
        model = model_match.group(0).strip() if model_match else "某车型"
        vehicles.append({
            "model": model,
            "vin": vin_clean,
            "collision": {"selected": True, "deductible": "500"},
            "comprehensive": {"selected": True, "deductible": "500"},
            "rental": {"selected": True, "limit": "30/30"},
            "roadside": {"selected": True},
        })
    return vehicles
