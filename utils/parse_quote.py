import os
import re
import fitz  # PyMuPDF
import boto3

def extract_text_from_file(file_path):
    """Extracts text from PDF or image using AWS Textract."""
    textract = boto3.client(
        "textract",
        aws_access_key_id=os.getenv("AWS_ACCESS_KEY_ID"),
        aws_secret_access_key=os.getenv("AWS_SECRET_ACCESS_KEY"),
        region_name=os.getenv("AWS_REGION")
    )

    ext = os.path.splitext(file_path)[-1].lower()

    if ext == ".pdf":
        with open(file_path, "rb") as f:
            bytes_data = f.read()
        response = textract.analyze_document(
            Document={"Bytes": bytes_data},
            FeatureTypes=["TABLES", "FORMS"]
        )
        text = " ".join([block["Text"] for block in response["Blocks"] if block["BlockType"] == "LINE"])
    elif ext in [".png", ".jpg", ".jpeg"]:
        with open(file_path, "rb") as f:
            image_bytes = f.read()
        response = textract.detect_document_text(Document={"Bytes": image_bytes})
        text = " ".join([block["Text"] for block in response["Blocks"] if block["BlockType"] == "LINE"])
    else:
        raise ValueError("Unsupported file format")

    return text

def parse_quote(file_path):
    raw_text = extract_text_from_file(file_path)

    data = {
        "company": None,
        "policy_term": None,
        "total_premium": None,
        "liability": {},
        "uninsured_motorist": {},
        "medical_payment": {},
        "personal_injury": {},
        "vehicles": []
    }

    # Detect company
    if "Progressive" in raw_text:
        data["company"] = "Progressive"
    elif "Travelers" in raw_text:
        data["company"] = "Travelers"
    else:
        data["company"] = "Unknown"

    # Total Premium
    total_match = re.search(r"Total .*?policy premium.*?\$([\d,\.]+)", raw_text, re.IGNORECASE)
    if total_match:
        data["total_premium"] = "$" + total_match.group(1)

    # Policy Term
    term_match = re.search(r"Total (\d+) month policy premium", raw_text, re.IGNORECASE)
    if term_match:
        data["policy_term"] = f"{term_match.group(1)}个月"

    # Liability
    bi_match = re.search(r"Bodily Injury Liability \$([\d,]+)/\$([\d,]+)", raw_text)
    pd_match = re.search(r"Property Damage Liability \$([\d,]+)", raw_text)
    if bi_match and pd_match:
        data["liability"] = {
            "selected": True,
            "bi_per_person": "$" + bi_match.group(1),
            "bi_per_accident": "$" + bi_match.group(2),
            "pd": "$" + pd_match.group(1)
        }
    else:
        data["liability"]["selected"] = False

    # Uninsured Motorist
    umbi_match = re.search(r"Uninsured.*?Bodily Injury \$([\d,]+)/\$([\d,]+)", raw_text)
    umpd_match = re.search(r"Uninsured.*?Property Damage \$([\d,]+)", raw_text)
    if umbi_match or umpd_match:
        data["uninsured_motorist"]["selected"] = True
        if umbi_match:
            data["uninsured_motorist"].update({
                "bi_per_person": "$" + umbi_match.group(1),
                "bi_per_accident": "$" + umbi_match.group(2)
            })
        if umpd_match:
            data["uninsured_motorist"].update({
                "pd": "$" + umpd_match.group(1),
                "deductible": "$250"
            })
    else:
        data["uninsured_motorist"]["selected"] = False

    # Medical Payments
    med_match = re.search(r"Medical Payments \$([\d,]+)", raw_text)
    if med_match:
        data["medical_payment"] = {
            "selected": True,
            "med": "$" + med_match.group(1)
        }
    else:
        data["medical_payment"]["selected"] = False

    # Personal Injury
    pip_match = re.search(r"Personal Injury Protection \$([\d,]+)", raw_text)
    if pip_match:
        data["personal_injury"] = {
            "selected": True,
            "pip": "$" + pip_match.group(1)
        }
    else:
        data["personal_injury"]["selected"] = False

    # Vehicles
    vehicle_blocks = re.findall(r"(\d{4} .*?VIN.*?[A-HJ-NPR-Z\d]{17})", raw_text, re.DOTALL)
    for block in vehicle_blocks:
        model_match = re.search(r"(\d{4}.*?)\s+VIN", block)
        vin_match = re.search(r"VIN[:\s]+([A-HJ-NPR-Z\d]{17})", block)
        if model_match and vin_match:
            vehicle = {
                "model": model_match.group(1).strip(),
                "vin": vin_match.group(1).strip(),
                "collision": "Collision" in block and "$" in block,
                "comprehensive": "Comprehensive" in block and "$" in block,
                "roadside": "Roadside" in block and "$" in block,
                "rental": "Rental" in block and "$" in block
            }
            data["vehicles"].append(vehicle)

    return data
