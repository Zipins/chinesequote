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
        print("\U0001F4E5 文件类型:", file_suffix)

        if file_suffix == "pdf":
            print("\U0001F50D PDF 文件，判断是否为文本型...")
            if is_pdf_text_based(pdf_bytes_for_fitz):
                print("✅ 是文本型 PDF，使用 PyMuPDF 提取文本")
                doc = fitz.open(stream=pdf_bytes_for_fitz, filetype="pdf")
                lines = []
                for page in doc:
                    lines.extend(page.get_text().splitlines())
                full_text = "\n".join(lines)
            else:
                print("\U0001F5BC 是扫描型 PDF，将 PDF 转为图片进行 OCR")
                images = pdf_to_images(file_bytes)
                if not images:
                    raise ValueError("PDF 转图片失败")
                all_text = []
                for idx, image_data in enumerate(images):
                    print(f"\U0001F4C4 正在处理第 {idx + 1} 页图像")
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
            print("\U0001F5BC️ 图片文件，直接 OCR")
            response = textract.detect_document_text(Document={"Bytes": file_bytes})
            lines = [block["Text"] for block in response["Blocks"] if block["BlockType"] == "LINE"]
            full_text = "\n".join(lines)
        else:
            print("❌ 不支持的文件类型")
            raise ValueError("不支持的文件格式")

        print("\U0001F4C4 OCR 文本提取完成，共", len(full_text), "字符")
        print("--- OCR 提取内容预览 ---\n", full_text[:1000], "\n--- END ---")

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

        print("✅ 字段提取完成:")
        for k, v in data.items():
            print(k, ":", v)

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
    known_companies = {
        "Progressive": "Progressive",
        "Travelers": "Travelers",
        "Allstate": "Allstate",
        "Geico": "Geico",
        "Liberty Mutual": "Liberty Mutual",
        "State Farm": "State Farm",
        "Safeco": "Safeco",
        "Nationwide": "Nationwide"
    }
    for name in known_companies:
        if name.lower() in text.lower():
            return known_companies[name]
    return "某保险公司"


def extract_total_premium(text):
    match = re.search(r"Total\s+\d+\s+month.*?[\r\n]+\$([\d,]+\.\d{2})", text, re.IGNORECASE)
    if match:
        return f"${match.group(1)}"
    return ""


def extract_policy_term(text):
    match = re.search(r"Total\s+(\d+)\s+month", text, re.IGNORECASE)
    if match:
        return f"{match.group(1)}个月"
    return ""


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
    result = {
        "selected": False, "bi_per_person": "", "bi_per_accident": "",
        "pd": "", "deductible": "250"
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
    pattern = re.compile(
        r"(?P<model>\d{4}\s+[A-Z0-9\s]+?)\nVIN[:\s]*(?P<vin>[A-HJ-NPR-Z0-9]{17})(.*?)(?=\n\d{4}\s+[A-Z]|$)",
        re.DOTALL
    )
    for match in pattern.finditer(text):
        model = match.group("model").strip()
        vin = match.group("vin").strip()
        block = match.group(0)
        vehicle = {
            "model": model,
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
        if vin in line:
            if re.search(r"\d{4} .*", line):
                return line
            for j in range(i - 1, max(i - 4, -1), -1):
                if re.search(r"\d{4} .*", lines[j]):
                    return lines[j]
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
    match = re.search(r"Rental.*?\$?([\d,]+)[^\d]+(\d+)[^\d]*day", text)
    if match:
        result["selected"] = True
        result["limit"] = f"{match.group(1)}/{match.group(2)}"
    return result


def extract_presence(text, keyword):
    return {"selected": keyword.lower() in text.lower()}
