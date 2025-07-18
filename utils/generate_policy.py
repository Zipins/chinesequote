from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import copy


def generate_policy_docx(doc: Document, data: dict):
    replace_placeholder_text(doc, "XXXXXXXXXXX", data.get("company", "某保险公司"))
    replace_placeholder_text(doc, "$XXXXXX/X个月", f"{data.get('total_premium', '$XXX')}/{data.get('policy_term', '6个月')}")

    # 写责任险
    write_checkbox_and_amount(doc, "Liability", data["liability"]["selected"])
    if data["liability"]["selected"]:
        replace_text_in_paragraphs(doc, "赔偿对方医疗费最高 $XXXX/人", f"赔偿对方医疗费最高 {data['liability']['bi_per_person']}/人")
        replace_text_in_paragraphs(doc, "赔偿对方医疗费总额最高$XXX", f"赔偿对方医疗费总额最高{data['liability']['bi_per_accident']}")
        replace_text_in_paragraphs(doc, "赔偿对方车辆和财产损失最多 $XXXX", f"赔偿对方车辆和财产损失最多 {data['liability']['pd']}")
    else:
        replace_text_in_paragraphs(doc, "赔偿对方医疗费最高 $XXXX/人", "没有选择该项目")
        replace_text_in_paragraphs(doc, "赔偿对方医疗费总额最高$XXX", "")
        replace_text_in_paragraphs(doc, "赔偿对方车辆和财产损失最多 $XXXX", "")

    # Uninsured Motorist
    write_checkbox_and_amount(doc, "Uninsured Motorist", data["uninsured_motorist"]["selected"])
    if data["uninsured_motorist"]["selected"]:
        replace_text_in_paragraphs(doc, "赔偿你和乘客医疗费$XXXX/人", f"赔偿你和乘客医疗费{data['uninsured_motorist']['bi_per_person']}/人")
        replace_text_in_paragraphs(doc, "一场事故最多赔偿医疗费$XXXX", f"一场事故最多赔偿医疗费{data['uninsured_motorist']['bi_per_accident']}")
        replace_text_in_paragraphs(doc, "赔偿自己车辆最多 $XXX(自付额$250)", f"赔偿自己车辆最多 {data['uninsured_motorist']['pd']}(自付额$250)")
    else:
        replace_text_in_paragraphs(doc, "赔偿你和乘客医疗费$XXXX/人", "没有选择该项目")
        replace_text_in_paragraphs(doc, "一场事故最多赔偿医疗费$XXXX", "")
        replace_text_in_paragraphs(doc, "赔偿自己车辆最多 $XXX(自付额$250)", "")

    # Medical Payment
    write_checkbox_and_amount(doc, "Medical Payment", data["medical_payment"]["selected"])
    if data["medical_payment"]["selected"]:
        replace_text_in_paragraphs(doc, "赔偿自己和自己车上乘客在事故中受伤的医疗费每人$XXX", f"赔偿自己和自己车上乘客在事故中受伤的医疗费每人{data['medical_payment']['med']}")
    else:
        replace_text_in_paragraphs(doc, "赔偿自己和自己车上乘客在事故中受伤的医疗费每人$XXX", "没有选择该项目")

    # Personal Injury
    write_checkbox_and_amount(doc, "Personal Injury", data["personal_injury"]["selected"])
    if data["personal_injury"]["selected"]:
        replace_text_in_paragraphs(doc, "赔偿自己和自己车上乘客在事故中受伤的医疗费，误工费和精神损失费每人$XXX", f"赔偿自己和自己车上乘客在事故中受伤的医疗费，误工费和精神损失费每人{data['personal_injury']['pip']}")
    else:
        replace_text_in_paragraphs(doc, "赔偿自己和自己车上乘客在事故中受伤的医疗费，误工费和精神损失费每人$XXX", "没有选择该项目")

    # 插入多辆车的保障信息
    insert_vehicle_sections(doc, data["vehicles"])


def insert_vehicle_sections(doc: Document, vehicles: list):
    from docx.shared import RGBColor
    from docx.oxml.ns import qn

    # 找到模板中“车辆保障:”段落
    insert_index = -1
    for i, para in enumerate(doc.paragraphs):
        if "车辆保障" in para.text:
            insert_index = i
            break
    if insert_index == -1:
        return

    # 模板中示例表格（用于复制格式）
    table_template = None
    for table in doc.tables:
        if "Collision" in table.cell(0, 0).text:
            table_template = table
            break
    if not table_template:
        return

    # 删除旧的表格和VIN段
    for i in reversed(range(insert_index + 1, len(doc.paragraphs))):
        if "VIN" in doc.paragraphs[i].text or "车辆" in doc.paragraphs[i].text:
            p = doc.paragraphs[i]._element
            p.getparent().remove(p)

    # 插入每辆车
    for idx, v in enumerate(vehicles):
        # 插入白色1pt空段
        blank_para = doc.paragraphs[insert_index].insert_paragraph_after("·")
        blank_para.runs[0].font.size = Pt(1)
        blank_para.runs[0].font.color.rgb = RGBColor(255, 255, 255)

        # 插入 VIN 信息段
        vin_para = blank_para.insert_paragraph_after(f"{v['model']}     VIN：{v['vin']}")
        vin_para.style.font.size = Pt(11)

        # 插入新表格
        new_table = copy.deepcopy(table_template._element)
        doc.paragraphs[insert_index]._element.addnext(new_table)
        table = Document().add_table(rows=5, cols=3)
        table._element = new_table
        fill_vehicle_table(table, v)


def fill_vehicle_table(table, vehicle):
    update_checkbox_cell(table.cell(1, 1), vehicle["collision"]["selected"])
    update_checkbox_cell(table.cell(2, 1), vehicle["comprehensive"]["selected"])
    update_checkbox_cell(table.cell(3, 1), vehicle["roadside"]["selected"])
    update_checkbox_cell(table.cell(4, 1), vehicle["rental"]["selected"])

    if vehicle["collision"]["selected"]:
        table.cell(1, 2).text = f"自付额${vehicle['collision']['deductible']}\n修车时自付额以内自己出，自付额以外的保险公司赔付"
    else:
        table.cell(1, 2).text = "没有选择该项目"

    if vehicle["comprehensive"]["selected"]:
        table.cell(2, 2).text = f"自付额${vehicle['comprehensive']['deductible']}修车时自付额以内自己出，自付额以外的保险公司赔付"
    else:
        table.cell(2, 2).text = "没有选择该项目"

    if vehicle["roadside"]["selected"]:
        table.cell(3, 2).text = "赔偿由于:机械故障,电瓶没电,钥匙被锁车内,燃油耗尽,轮胎没气造成车辆不可行驶时的免费拖车，免费充电，免费开锁服务"
    else:
        table.cell(3, 2).text = "没有选择该项目"

    if vehicle["rental"]["selected"]:
        table.cell(4, 2).text = "每天$30最多30天"
    else:
        table.cell(4, 2).text = "没有选择该项目"


def update_checkbox_cell(cell, selected):
    cell.text = "✅" if selected else "❌"
    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = cell.paragraphs[0].runs[0]
    run.font.size = Pt(16)


def replace_placeholder_text(doc, placeholder, replacement):
    for paragraph in doc.paragraphs:
        if placeholder in paragraph.text:
            paragraph.text = paragraph.text.replace(placeholder, replacement)


def replace_text_in_paragraphs(doc, old, new):
    for paragraph in doc.paragraphs:
        if old in paragraph.text:
            paragraph.text = paragraph.text.replace(old, new)


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
