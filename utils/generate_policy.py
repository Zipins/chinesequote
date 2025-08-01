from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from copy import deepcopy

def generate_policy_docx(doc: Document, data: dict):
    # 替换价格信息
    price_info = f"{data.get('total_premium', '$XXX')}/{data.get('policy_term', '6个月')}，一次性付款"
    replace_placeholder_text(doc, "{{PRICE_INFO}}", price_info)

    # 责任险
    write_checkbox_and_amount(doc, "Liability", data["liability"]["selected"])
    if data["liability"]["selected"]:
        replace_placeholder_text(doc, "{{LIAB_BI_PP}}", data['liability']['bi_per_person'])
        replace_placeholder_text(doc, "{{LIAB_BI_PA}}", data['liability']['bi_per_accident'])
        replace_placeholder_text(doc, "{{LIAB_PD}}", data['liability']['pd'])
    else:
        clear_liability_section(doc)

    # 无保险驾驶者
    write_checkbox_and_amount(doc, "Uninsured Motorist", data["uninsured_motorist"]["selected"])
    if data["uninsured_motorist"]["selected"]:
        if data["uninsured_motorist"].get("bi_per_person"):
            replace_placeholder_text(doc, "{{UNINS_BI_PP}}", data['uninsured_motorist']['bi_per_person'])
        if data["uninsured_motorist"].get("bi_per_accident"):
            replace_placeholder_text(doc, "{{UNINS_BI_PA}}", data['uninsured_motorist']['bi_per_accident'])
        if data["uninsured_motorist"].get("pd"):
            replace_placeholder_text(doc, "{{UNINS_PD}}", data['uninsured_motorist']['pd'])
    else:
        clear_uninsured_section(doc)

    # Medical Payment
    write_checkbox_and_amount(doc, "Medical Payment", data["medical_payment"]["selected"])
    if data["medical_payment"]["selected"]:
        replace_placeholder_text(doc, "{{MED}}", data['medical_payment']['med'])
    else:
        replace_placeholder_text(doc, "{{MED}}", "没有选择该项目")

    # Personal Injury
    write_checkbox_and_amount(doc, "Personal Injury", data["personal_injury"]["selected"])
    if data["personal_injury"]["selected"]:
        replace_placeholder_text(doc, "{{PIP}}", data['personal_injury']['pip'])
    else:
        replace_placeholder_text(doc, "{{PIP}}", "没有选择该项目")

    # 插入车辆保障
    insert_vehicle_section(doc, data.get("vehicles", []))

def clear_liability_section(doc):
    replace_placeholder_text(doc, "{{LIAB_BI_PP}}", "没有选择该项目")
    replace_placeholder_text(doc, "{{LIAB_BI_PA}}", "")
    replace_placeholder_text(doc, "{{LIAB_PD}}", "")

def clear_uninsured_section(doc):
    replace_placeholder_text(doc, "{{UNINS_BI_PP}}", "没有选择该项目")
    replace_placeholder_text(doc, "{{UNINS_BI_PA}}", "")
    replace_placeholder_text(doc, "{{UNINS_PD}}", "")

def replace_placeholder_text(doc, placeholder, replacement):
    for paragraph in doc.paragraphs:
        if placeholder in paragraph.text:
            paragraph.text = paragraph.text.replace(placeholder, replacement)

def write_checkbox_and_amount(doc, keyword, selected):
    symbol = "✅" if selected else "❌"
    for table in doc.tables:
        for row in table.rows:
            if keyword in row.cells[0].text:
                cell = row.cells[1]
                cell.text = symbol
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = cell.paragraphs[0].runs[0]
                run.font.size = Pt(16)

# insert_vehicle_section 保持你现有版本或我之前给你的完整版即可
