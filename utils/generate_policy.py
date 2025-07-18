from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import streamlit as st  # 用于调试输出金额字段


def generate_policy_docx(doc: Document, data: dict):
    replace_placeholder_text(doc, "XXXXXXXXXXX", data.get("company", "某保险公司"))
    replace_placeholder_text(doc, "$XXXXXX/X个月", f"{data.get('total_premium', '$XXX')}/{data.get('policy_term', '6个月')}")

    # 写入责任险
    write_checkbox_and_amount(doc, "Liability", data["liability"]["selected"])
    if data["liability"]["selected"]:
        replace_text_in_paragraphs(doc, "赔偿对方医疗费最高$XXXX/人", f"赔偿对方医疗费最高{data['liability']['bi_per_person']}/人")
        replace_text_in_paragraphs(doc, "赔偿对方医疗费总额最高$XXX", f"赔偿对方医疗费总额最高{data['liability']['bi_per_accident']}")
        replace_text_in_paragraphs(doc, "赔偿对方车辆和财产损失最多$XXXX", f"赔偿对方车辆和财产损失最多{data['liability']['pd']}")
    else:
        replace_text_in_paragraphs(doc, "赔偿对方医疗费最高$XXXX/人", "没有选择该项目")
        replace_text_in_paragraphs(doc, "赔偿对方医疗费总额最高$XXX", "")
        replace_text_in_paragraphs(doc, "赔偿对方车辆和财产损失最多$XXXX", "")

    # Uninsured Motorist
    write_checkbox_and_amount(doc, "Uninsured Motorist", data["uninsured_motorist"]["selected"])
    if data["uninsured_motorist"]["selected"]:
        replace_text_in_paragraphs(doc, "赔偿你和乘客医疗费$XXXX/人", f"赔偿你和乘客医疗费{data['uninsured_motorist']['bi_per_person']}/人")
        replace_text_in_paragraphs(doc, "一场事故最多赔偿医疗费$XXXX", f"一场事故最多赔偿医疗费{data['uninsured_motorist']['bi_per_accident']}")
        replace_text_in_paragraphs(doc, "赔偿自己车辆最多$XXX(自付额$250)", f"赔偿自己车辆最多{data['uninsured_motorist']['pd']}(自付额$250)")
    else:
        replace_text_in_paragraphs(doc, "赔偿你和乘客医疗费$XXXX/人", "没有选择该项目")
        replace_text_in_paragraphs(doc, "一场事故最多赔偿医疗费$XXXX", "")
        replace_text_in_paragraphs(doc, "赔偿自己车辆最多$XXX(自付额$250)", "")

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

    # 插入车辆保障表格
    insert_vehicle_section(doc, data.get("vehicles", []))

    # 调试用：在 Streamlit 页面上展示所有包含金额字段的段落
    print_all_paragraphs_with_dollar(doc)


def insert_vehicle_section(doc, vehicles):
    from copy import deepcopy

    marker_idx = -1
    for i, p in enumerate(doc.paragraphs):
        if "车辆保障:" in p.text:
            marker_idx = i
            break
    if marker_idx == -1:
        return

    marker = doc.paragraphs[marker_idx]._element

    # 清除原有表格和 VIN 信息
    next_el = marker.getnext()
    while next_el is not None and (next_el.tag.endswith("p") or next_el.tag.endswith("tbl")):
        to_remove = next_el
        next_el = next_el.getnext()
        marker.getparent().remove(to_remove)

    doc_paragraph = doc.paragraphs[marker_idx]

    for vehicle in vehicles:
        spacer_p = doc_paragraph.insert_paragraph_before("·")
        spacer_p.runs[0].font.size = Pt(1)
        spacer_p.runs[0].font.color.rgb = RGBColor(255, 255, 255)

        vin_para = doc_paragraph.insert_paragraph_before(f"{vehicle['model']}     VIN：{vehicle['vin']}")
        vin_para.runs[0].font.size = Pt(12)
        vin_para.runs[0].bold = True

        table = doc.add_table(rows=5, cols=3)
        table.style = "Table Grid"
        table.autofit = False
        table.allow_autofit = False

        headers = ["保险项目", "是否选择", "保额 / 说明"]
        chinese_rows = ["碰撞险", "车子损毁险", "道路救援", "租车报销"]
        for j, header in enumerate(headers):
            table.cell(0, j).text = header
        for i in range(4):
            table.cell(i + 1, 0).text = chinese_rows[i]

        fill_vehicle_table(table, vehicle)
        marker.addnext(table._element)
        marker = table._element


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
        table.cell(2, 2).text = f"自付额${vehicle['comprehensive']['deductible']}\n修车时自付额以内自己出，自付额以外的保险公司赔付"
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
        full_text = "".join(run.text for run in paragraph.runs)
        if old in full_text:
            new_full_text = full_text.replace(old, new)
            # 清空现有 runs
            for run in paragraph.runs:
                run.text = ""
            # 只用一个 run 写入新内容
            paragraph.runs[0].text = new_full_text


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


def print_all_paragraphs_with_dollar(doc):
    st.subheader("🔍 模板中包含金额字段（$）的段落")
    for i, para in enumerate(doc.paragraphs):
        if "$" in para.text or "￥" in para.text:
            st.markdown(f"**段落 {i+1}:** `{para.text.strip()}`")
