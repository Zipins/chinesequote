import re
import fitz  # PyMuPDF
import boto3
from io import BytesIO
from PIL import Image
from pdf2image import convert_from_bytes

def format_dollar(amount):
    if not amount:
        return ""
    return "$" + "{:,}".format(int(amount.replace("$", "").replace(",", "").strip()))

def extract_text_from_file(file_bytes, filename):
    textract = boto3.client(
       "textract",
       region_name=os.getenv("AWS_REGION"),
       aws_access_key_id=os.getenv("AWS_ACCESS_KEY_ID"),
       aws_secret_access_key=os.getenv("AWS_SECRET_ACCESS_KEY")
)
    
    if filename.lower().endswith(".pdf"):
        with fitz.open(stream=file_bytes, filetype="pdf") as doc:
            is_text_pdf = any(page.get_text().strip() for page in doc)
        
        if is_text_pdf:
            text = ""
            for page in fitz.open(stream=file_bytes, filetype="pdf"):
                text += page.get_text()
            return text
        else:
            images = convert_from_bytes(file_bytes)
            text = ""
            for image in images:
                buffered = BytesIO()
                image.save(buffered, format="PNG")
                response = textract.detect_document_text(
                    Document={'Bytes': buffered.getvalue()}
                )
                for item in response["Blocks"]:
                    if item["BlockType"] == "LINE":
                        text += item["Text"] + "\n"
            return text

    elif filename.lower().endswith(('.png', '.jpg', '.jpeg')):
        response = textract.detect_document_text(
            Document={'Bytes': file_bytes}
        )
        text = ""
        for item in response["Blocks"]:
            if item["BlockType"] == "LINE":
                text += item["Text"] + "\n"
        return text

    else:
        raise ValueError("Unsupported file format")

def parse_quote_from_text(text):
    data = {
        "company": "Progressive" if "progressive" in text.lower() else "",
        "total_premium": None,
        "policy_term": None,
        "liability": {"selected": False, "bi_per_person": "", "bi_per_accident": "", "pd": ""},
        "uninsured motorist": {"selected": False, "bi_per_person": "", "bi_per_accident": "", "pd": "", "deductible": ""},
        "medical payment": {"selected": False},
        "personal injury": {"selected": False, "pip": None},
        "vehicles": []
    }

    # total premium
    premium_match = re.search(r"Total \d+ month policy premium.*?\$([\d,]+\.\d{2})", text, re.IGNORECASE)
    if premium_match:
        data["total_premium"] = "$" + premium_match.group(1)

    term_match = re.search(r"Total (\d+) month policy premium", text)
    if term_match:
        data["policy_term"] = term_match.group(1) + "个月"

    # liability
    if "Liability to Others" in text:
        data["liability"]["selected"] = True
        bi_match = re.search(r"Bodily Injury Liability \$(\d{1,3}(?:,\d{3})*) each person/\$(\d{1,3}(?:,\d{3})*) each accident", text)
        pd_match = re.search(r"Property Damage Liability \$(\d{1,3}(?:,\d{3})*) each accident", text)
        if bi_match:
            data["liability"]["bi_per_person"] = "$" + bi_match.group(1)
            data["liability"]["bi_per_accident"] = "$" + bi_match.group(2)
        if pd_match:
            data["liability"]["pd"] = "$" + pd_match.group(1)

    # uninsured motorist
    if "Uninsured/Underinsured Motorist" in text:
        umbi_match = re.search(r"Uninsured/Underinsured Motorist Bodily Injury \$(\d{1,3}(?:,\d{3})*) each person/\$(\d{1,3}(?:,\d{3})*) each accident", text)
        umpd_match = re.search(r"Uninsured/Underinsured Motorist Property Damage \$(\d{1,3}(?:,\d{3})*) each accident", text)
        if umbi_match or umpd_match:
            data["uninsured motorist"]["selected"] = True
        if umbi_match:
            data["uninsured motorist"]["bi_per_person"] = "$" + umbi_match.group(1)
            data["uninsured motorist"]["bi_per_accident"] = "$" + umbi_match.group(2)
        if umpd_match:
            data["uninsured motorist"]["pd"] = "$" + umpd_match.group(1)
            data["uninsured motorist"]["deductible"] = "$250"  # 默认 Progressive 自付额固定

    # medical payment
    if re.search(r"Medical Payments.*\$(\d{1,3}(?:,\d{3})*)", text):
        data["medical payment"]["selected"] = True

    # personal injury
    pip_match = re.search(r"Personal Injury Protection \$?(\d{1,3}(?:,\d{3})*)", text)
    if pip_match:
        data["personal injury"]["selected"] = True
        data["personal injury"]["pip"] = int(pip_match.group(1).replace(",", ""))

    # vehicles
    vehicle_blocks = re.findall(r"(\d{4} [A-Z0-9 \-]+?)\n?VIN[:： ]([A-HJ-NPR-Z0-9]{11,})[\s\S]*?Premium: \$[\d,.]+", text)
    for model, vin in vehicle_blocks:
        vehicle = {
            "model": model.strip(),
            "vin": vin.strip(),
            "collision": "✅" if re.search(vin + r"[\s\S]{0,100}?Collision", text) else "❌",
            "comprehensive": "✅" if re.search(vin + r"[\s\S]{0,100}?Comprehensive", text) else "❌",
            "rental": "✅" if re.search(vin + r"[\s\S]{0,100}?Rental", text) else "❌",
            "roadside": "✅" if re.search(vin + r"[\s\S]{0,100}?Roadside", text) else "❌"
        }
        data["vehicles"].append(vehicle)

    return data
