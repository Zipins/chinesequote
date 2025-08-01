
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from copy import deepcopy


def generate_policy_docx(doc: Document, data: dict):
    replace_paragraph_text(doc, "XXXXXXXXXXX", data.get("company", "某保险公司"))
    replace_paragraph_text(doc, "$XXXXXX/X个月", f"{data.get('total_premium', '$XXX')}/{data.get('policy_term', '6个月')}")

    # 写入责任险
    write_checkbox_and_amount(doc, "Liability", data["liability"]["selected"])
    if data["liability"]["selected"]:
        replace_paragraph_text(doc, "赔偿对方医疗费最高$XXXX/人", f"赔偿对方医疗费最高{data['liability']['bi_per_person']}/人")
        replace_paragraph_text(doc, "赔偿对方医疗费总额最高$XXX", f"赔偿对方医疗费总额最高{data['liability']['bi_per_accident']}")
        replace_paragraph_text(doc, "赔偿对方车辆和财产损失最多$XXXX", f"赔偿对方车辆和财产损失最多{data['liability']['pd']}")
    else:
        replace_paragraph_text(doc, "赔偿对方医疗费最高$XXXX/人", "没有选择该项目")
        replace_paragraph_text(doc, "赔偿对方医疗费总额最高$XXX", "")
        replace_paragraph_text(doc, "赔偿对方车辆和财产损失最多$XXXX", "")

    # Uninsured Motorist
    write_checkbox_and_amount(doc, "Uninsured Motorist", data["uninsured_motorist"]["selected"])
    if data["uninsured_motorist"]["selected"]:
        um = data["uninsured_motorist"]
        desc = []
        if um.get("bi_per_person") and um.get("bi_per_accident"):
            desc.append(f"赔偿你和乘客医疗费{um['bi_per_person']}/人")
            desc.append(f"一场事故最多赔偿医疗费{um['bi_per_accident']}")
        if um.get("pd"):
            desc.append(f"赔偿自己车辆{um['pd']}(自付额$250)")
        full_desc = "\n".join(desc) if desc else "没有选择该项目"
        replace_paragraph_text(doc, "赔偿你和乘客医疗费$XXXX/人", full_desc)
    else:
        replace_paragraph_text(doc, "赔偿你和乘客医疗费$XXXX/人", "没有选择该项目")
        replace_paragraph_text(doc, "一场事故最多赔偿医疗费$XXXX", "")
        replace_paragraph_text(doc, "赔偿自己车辆最多$XXX(自付额$250)", "")

    # Medical Payment
    write_checkbox_and_amount(doc, "Medical Payment", data["medical_payment"]["selected"])
    if data["medical_payment"]["selected"]:
        replace_paragraph_text(doc, "赔偿自己和自己车上乘客在事故中受伤的医疗费每人$XXX",
                               f"赔偿自己和自己车上乘客在事故中受伤的医疗费每人{data['medical_payment']['med']}")
    else:
        replace_paragraph_text(doc, "赔偿自己和自己车上乘客在事故中受伤的医疗费每人$XXX", "没有选择该项目")

    # Personal Injury
    write_checkbox_and_amount(doc, "Personal Injury", data["personal_injury"]["selected"])
    if data["personal_injury"]["selected"]:
        replace_paragraph_text(doc, "赔偿自己和自己车上乘客在事故中受伤的医疗费，误工费和精神损失费每人$XXX",
                               f"赔偿自己和自己车上乘客在事故中受伤的医疗费，误工费和精神损失费每人{data['personal_injury']['pip']}")
    else:
        replace_paragraph_text(doc, "赔偿自己和自己车上乘客在事故中受伤的医疗费，误工费和精神损失费每人$XXX", "没有选择该项目")


def replace_paragraph_text(doc: Document, old_text: str, new_text: str):
    for paragraph in doc.paragraphs:
        if old_text in paragraph.text:
            for run in paragraph.runs:
                run.text = ""
            paragraph.add_run(new_text)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if old_text in paragraph.text:
                        for run in paragraph.runs:
                            run.text = ""
                        paragraph.add_run(new_text)


def write_checkbox_and_amount(doc, keyword, selected):
    symbol = "✅" if selected else "❌"
    for table in doc.tables:
        for row in table.rows:
            if keyword in row.cells[0].text:
                cell = row.cells[1]
                cell.text = symbol
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.runs[0] if p.runs else p.add_run()
                run.font.size = Pt(16)
