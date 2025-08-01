from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from copy import deepcopy

def generate_policy_docx(doc: Document, data: dict):
    # 替换价格信息
    price_info = f"{data.get('total_premium', '$XXX')}/{data.get('policy_term', '6个月')}，一次性付款"
    replace_placeholder_text(doc, "{{PRICE_INFO}}", price_info)
    replace_placeholder_text(doc, "{{COMPANY}}", data.get("company", "某保险公司"))

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

def insert_vehicle_section(doc: Document, vehicles: list):
    if not vehicles:
        return

    # 找到“车辆保障:”段落
    marker_idx = -1
    for i, p in enumerate(doc.paragraphs):
        if "车辆保障:" in p.text:
            marker_idx = i
            break
    if marker_idx == -1:
        return

    marker_p = doc.paragraphs[marker_idx]
    marker_el = marker_p._element

    # 清除后续旧表格和 VIN 信息
    next_el = marker_el.getnext()
    while next_el is not None and (next_el.tag.endswith("p") or next_el.tag.endswith("tbl")):
        to_remove = next_el
        next_el = next_el.getnext()
        marker_el.getparent().remove(to_remove)

    # 查找模板中第一个完整车辆保障表格作为复制模板
    vehicle_table_template = None
    for tbl in doc.tables:
        if "Collision" in tbl.cell(1, 0).text and "租车报销" in tbl.cell(4, 0).text:
            vehicle_table_template = tbl
            break
    if not vehicle_table_template:
        return

    for vehicle in vehicles:
        # 插入视觉空行
        spacer_p = marker_p.insert_paragraph_after("·")
        spacer_p.runs[0].font.size = Pt(1)
        spacer_p.runs[0].font.color.rgb = RGBColor(255, 255, 255)

        # 插入 VIN 信息
        vin_text = f"{vehicle['model']}     VIN：{vehicle['vin']}"
        vin_p = spacer_p.insert_paragraph_after(vin_text)
        vin_p.runs[0].font.size = Pt(12)
        vin_p.runs[0].bold = True

        # 插入复制表格
        new_table = deepcopy(vehicle_table_template._element)
        vin_p._element.addnext(new_table)
        new_table_obj = vin_p._element.getnext()

        fill_vehicle_table(doc, new_table_obj, vehicle)

        marker_p = doc.paragraphs[-1]

def fill_vehicle_table(doc: Document, table_el, vehicle: dict):
    tbl = None
    for t in doc.tables:
        if t._element == table_el:
            tbl = t
            break
    if not tbl:
        return

    update_checkbox_cell(tbl.cell(1, 1), vehicle["collision"]["selected"])
    update_checkbox_cell(tbl.cell(2, 1), vehicle["comprehensive"]["selected"])
    update_checkbox_cell(tbl.cell(3, 1), vehicle["roadside"]["selected"])
    update_checkbox_cell(tbl.cell(4, 1), vehicle["rental"]["selected"])

    if vehicle["collision"]["selected"]:
        tbl.cell(1, 2).text = f"自付额${vehicle['collision']['deductible']}\n修车时自付额以内自己出，自付额以外的保险公司赔付"
    else:
        tbl.cell(1, 2).text = "没有选择该项目"

    if vehicle["comprehensive"]["selected"]:
        tbl.cell(2, 2).text = f"自付额${vehicle['comprehensive']['deductible']}\n修车时自付额以内自己出，自付额以外的保险公司赔付"
    else:
        tbl.cell(2, 2).text = "没有选择该项目"

    if vehicle["roadside"]["selected"]:
        tbl.cell(3, 2).text = "赔偿由于:机械故障,电瓶没电,钥匙被锁车内,燃油耗尽,轮胎没气造成车辆不可行驶时的免费拖车，免费充电，免费开锁服务"
    else:
        tbl.cell(3, 2).text = "没有选择该项目"

    if vehicle["rental"]["selected"]:
        tbl.cell(4, 2).text = "每天$30 最多30天"
    else:
        tbl.cell(4, 2).text = "没有选择该项目"

def update_checkbox_cell(cell, selected):
    cell.text = "✅" if selected else "❌"
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.runs[0]
    run.font.size = Pt(16)
