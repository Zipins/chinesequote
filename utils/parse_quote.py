import boto3
import re
import io
import fitz  # PyMuPDF
from PIL import Image
import traceback

def extract_quote_data(file, return_raw_text=False):
    file_bytes = file.read()
    pdf_bytes_for_fitz = file_bytes  # ✅ 用于判断是否为文本型 PDF
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
        print("📥 文件类型:", file_suffix)

        if file_suffix == "pdf":
            print("🔍 PDF 文件，判断是否为文本型...")
            if is_pdf_text_based(pdf_bytes_for_fitz):
                print("✅ 是文本型 PDF，使用 detect_document_text")
                response = textract.detect_document_text(Document={"Bytes": file_bytes})
                lines = [block["Text"] for block in response["Blocks"] if block["BlockType"] == "LINE"]
                full_text = "\n".join(lines)
            else:
                print("📷 是扫描型 PDF，将 PDF 转为图片进行处理")
                images = pdf_to_images(file_bytes)
                if not images:
                    raise ValueError("PDF 转图片失败")
                all_text = []
                for idx, image_data in enumerate(images):
                    print(f"📄 正在处理第 {idx + 1} 页图像")
                    image_bytes_io = io.BytesIO(image_data)
                    img_response = textract.detect_document_text(Document={"Bytes": image_bytes_io.read()})
                    for block in img_response["Blocks"]:
                        if block["BlockType"] == "LINE":
                            all_text.append(block["Text"])
                full_text = "\n".join(all_text)
        elif file_suffix in ["jpg", "jpeg", "png"]:
            print("🖼️ 图片文件，直接 OCR")
            response = textract.detect_document_text(Document={"Bytes": file_bytes})
            lines = [block["Text"] for block in response["Blocks"] if block["BlockType"] == "LINE"]
            full_text = "\n".join(lines)
        else:
            print("❌ 不支持的文件类型")
            raise ValueError("不支持的文件格式")

        print("📄 OCR 文本提取完成，共", len(full_text), "字符")
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

        print("✅ 字段提取完成")
        return (data, full_text) if return_raw_text else data

    except Exception as e:
        print("❌ 异常发生：", str(e))
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
