import boto3
import re
import io

def extract_quote_data(file, return_raw_text=False):
    # 将文件读取为字节
    file_bytes = file.read()

    # 调用 Textract
    client = boto3.client("textract")
    response = client.detect_document_text(Document={"Bytes": file_bytes})

    lines = []
    for block in response["Blocks"]:
        if block["BlockType"] == "LINE":
            lines.append(block["Text"])
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

# 以下是字段提取函数

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

    # 分割每辆车的信息块
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
