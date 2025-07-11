# generate_policy.py

import os
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

def generate_policy_docx(data, output_path):
    template_path = "template/\u4fdd\u5355\u8303\u4f8b.docx"
    doc = Document(template_path)

    def replace_placeholder(paragraphs, placeholder, value):
        for p in paragraphs:
            if placeholder in p.text:
                inline = p.runs
                for i in range(len(inline)):
                    if placeholder in inline[i].text:
                        inline[i].text = inline[i].text.replace(placeholder, value)

    replace_placeholder(doc.paragraphs, "{{insurance_company}}", data.get("company", ""))
    replace_placeholder(doc.paragraphs, "{{total_premium}}", data.get("total_premium", ""))
    replace_placeholder(doc.paragraphs, "{{policy_term}}", data.get("policy_term", ""))

    def write_coverage(table, item, selected, description):
        for row in table.rows:
            if item in row.cells[0].text:
                row.cells[1].text = "\u2705" if selected else "\u274c"
                row.cells[1].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                run = row.cells[1].paragraphs[0].runs[0]
                run.font.size = Pt(16)
                row.cells[2].text = description
                row.cells[2].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    tables = doc.tables
    write_coverage(tables[0], "\u8d23\u4efb\u9669", True, data["liability"]["description"])
    write_coverage(tables[0], "\u65e0\u4fdd\u9669\u9a7e\u9a76\u8005", data["uninsured_motorist"].get("selected", False), data["uninsured_motorist"].get("description", "\u6ca1\u6709\u9009\u62e9\u8be5\u9879\u76ee"))
    write_coverage(tables[0], "\u533b\u7597\u8d39\u7528", data["medical_payment"].get("selected", False), data["medical_payment"].get("description", "\u6ca1\u6709\u9009\u62e9\u8be5\u9879\u76ee"))
    write_coverage(tables[0], "\u4eba\u8eab\u635f\u5931", data["personal_injury"].get("selected", False), data["personal_injury"].get("description", "\u6ca1\u6709\u9009\u62e9\u8be5\u9879\u76ee"))

    def insert_vehicle_table(doc, vehicle):
        doc.add_paragraph("\u00b7", style=None).runs[0].font.size = Pt(1)
        vin_text = f"{vehicle['model']}     VIN\uff1a{vehicle['vin']}"
        doc.add_paragraph(vin_text)
        table = doc.add_table(rows=5, cols=3)
        table.style = doc.styles["Table Grid"]
        items = [
            ("Collision", "\u8f66\u8f86\u6467\u6495\u9669"),
            ("Comprehensive", "\u975e\u78b0\u6495\u9669"),
            ("Roadside", "\u62a4\u7406\u6551\u63f4"),
            ("Rental", "\u79df\u8f66\u8865\u507f")
        ]
        for i, (key, label) in enumerate(items):
            table.cell(i, 0).text = label
            table.cell(i, 1).text = "\u2705" if vehicle.get(key.lower()) else "\u274c"
            table.cell(i, 2).text = f"\u9009\u62e9{label}\u4fdd\u969c\u9879\u76ee\u3002"

    doc.add_paragraph("\u00b7", style=None).runs[0].font.size = Pt(1)
    doc.add_paragraph("\u8f66\u8f86\u4fdd\u969c\uff1a", style=None).runs[0].font.size = Pt(14)
    for v in data.get("vehicles", []):
        insert_vehicle_table(doc, v)

    doc.save(output_path)
