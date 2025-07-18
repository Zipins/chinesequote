import streamlit as st
from docx import Document
import tempfile
from utils.parse_quote import extract_quote_data
from utils.generate_policy import generate_policy_docx  # 确保此函数已支持多车辆
import os

TEMPLATE_PATH = "template/保单范例.docx"

st.set_page_config(page_title="中文保单生成器", layout="centered")

st.title("📄 中文保单生成器")
st.markdown("上传保险报价 PDF 或图片，系统将自动生成中文保单解释文档。")

uploaded_file = st.file_uploader("请上传报价文件（PDF / JPG / PNG）", type=["pdf", "jpg", "jpeg", "png"])

if uploaded_file:
    with st.spinner("正在识别报价内容，请稍候..."):
        try:
            # 1. 从 Textract 提取数据和 OCR 文本
            data, ocr_text = extract_quote_data(uploaded_file, return_raw_text=True)

            # 2. 加载 Word 模板
            template_doc = Document(TEMPLATE_PATH)

            # 3. 生成中文保单（支持多车辆）
            generate_policy_docx(template_doc, data)

            # 4. 保存为临时文件
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                output_path = tmp.name
                template_doc.save(output_path)

            # 5. 成功提示 + 下载按钮
            st.success("✅ 保单生成成功！")
            st.download_button("📥 下载生成的中文保单", data=open(output_path, "rb"), file_name="中文保单.docx")

            # 6. 显示字段提取结果（调试或验证用）
            st.subheader("📋 提取字段预览")
            st.json(data)

            # 7. 展示 OCR 原文（可折叠）
            with st.expander("🧾 OCR 原始文本（调试用）"):
                st.text(ocr_text)

        except Exception as e:
            st.error(f"❌ 出错了：{e}")
