# generate_policy.py
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import os

def generate_policy_docx(data, template_path, output_path):
    doc = Document(template_path)

    def replace_text(placeholder, value):
        for p in doc.paragraphs:
            if placeholder in p.text:
                for run in p.runs:
                    if placeholder in run.text:
                        run.text = run.text.replace(placeholder, value)

    replace_text("{{company}}", data.get("company", ""))
    replace_text("{{total_premium}}", data.get("total_premium", ""))
    replace_text("{{policy_term}}", data.get("policy_term", ""))

    # 责任险部分
    liability_info = data.get("liability", {})
    for table in doc.tables:
        for row in table.rows:
            cells = row.cells
            if len(cells) >= 3:
                if "责任险（对第三方）" in cells[0].text:
                    selected_cell = cells[1].paragraphs[0]
                    selected_cell.clear()
                    selected_cell.add_run("✅" if liability_info.get("selected") else "❌").font.size = Pt(16)
                    amount_cell = cells[2].paragraphs[0]
                    if liability_info.get("selected"):
                        desc = (
                            f"赔偿对方医疗费最高 {liability_info.get('bi_per_person', '')}/人\n"
                            f"赔偿对方医疗费总额最高{liability_info.get('bi_per_accident', '')}\n"
                            f"赔偿对方车辆和财产损失最多{liability_info.get('pd', '')}"
                        )
                    else:
                        desc = "没有选择该项目"
                    amount_cell.clear()
                    amount_cell.add_run(desc)

    # TODO: Add rest of the coverage + vehicles + formatting logic
    doc.save(output_path)
