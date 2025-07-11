import boto3
import re
import io
from PyPDF2 import PdfReader
from PIL import Image
from pdf2image import convert_from_bytes
import tempfile

def extract_quote_data(file, return_raw_text=False):
    file_bytes = file.read()
    filename = file.name.lower()

    full_text = ""

    if filename.endswith(".pdf"):
        try:
            reader = PdfReader(io.BytesIO(file_bytes))
            full_text = "\n".join([page.extract_text() or "" for page in reader.pages])
            if not full_text.strip():
                raise ValueError("Empty text, fallback to Textract")
        except:
            images = convert_from_bytes(file_bytes)
            full_text = textract_image_sequence(images)
    elif filename.endswith(('.png', '.jpg', '.jpeg')):
        image = Image.open(io.BytesIO(file_bytes))
        full_text = textract_image(image)
    else:
        raise ValueError("Unsupported file format")

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

def textract_image(image):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
        image.save(tmp.name, format="PNG")
        with open(tmp.name, "rb") as f:
            image_bytes = f.read()

    client = boto3.client("textract", region_name="us-east-1")
    response = client.detect_document_text(Document={"Bytes": image_bytes})
    lines = [block["Text"] for block in response["Blocks"] if block["BlockType"] == "LINE"]
    return "\n".join(lines)

def textract_image_sequence(images):
    return "\n".join(textract_image(img) for img in images)

# 字段提取函数

def extract_company_name(text):
    if "Progressive" in text:
        return "Progressive"
    if "Travelers" in text:
        return "Travelers"
    return "某保险公司"

def extract_total_premium(text):
    match = re.search(r"Total\s+\d+\s+month.*?\$([\d,]+\.\d{2})", text, re.IGNORECASE)
    return f"${match.group(1)}" if match else ""

def extract_policy_term(text):
    match = re.search(r"Total\s+(\d+)\s+month", text, re.IGNORECASE)
    return f"{match.group(1)}个月" if match else ""

def extract_liability(text):
    result = {"selected": False, "bi_per_person": "", "bi_per_accident": "", "pd": ""}
    bi_match = re.search(r"Bodily Injury Liability\s*\$([\d,]+)[^\d]+([\d,]+)", text)
    pd_match = re.search(r"Property Damage Liability\s*\$([\d,]+)", text)
    if bi_match and pd_match:
        result["selected"] = True
        result["bi_per_person"] = f"${bi_match.group(1)}"
        result["bi_per_accident"] = f"${bi_match.group(2)}"
        result["pd"] = f"${pd_match.group(1)}"
    return result

def extract_uninsured_motorist(text):
    result = {"selected": False, "bi_per_person": "", "bi_per_accident": "", "pd": "", "deductible": "250"}
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
    blocks = re.split(r"(?=VIN[:\s])", text)
    for block in blocks:
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
