import streamlit as st
from docx import Document
import tempfile
from utils.parse_quote import extract_quote_data
from utils.generate_policy import generate_policy_docx
import os

TEMPLATE_PATH = "template/保单范例.docx"

st.set_page_config(page_title="中文保单生成器", layout="centered")

st.title("📄 中文保单生成器")
st.markdown("上传保险报价 PDF 或图片，系统将自动生成中文保单解释文档。")

uploaded_file = st.file_uploader("请上传报价文件（PDF / JPG / PNG）", type=["pdf", "jpg", "jpeg", "png"])

if uploaded_file:
    with st.spinner("正在识别报价内容，请稍候..."):
        try:
            # 从 Textract 提取数据
            data, ocr_text = extract_quote_data(uploaded_file, return_raw_text=True)

            # 打开模板文档
            template_doc = Document(TEMPLATE_PATH)

            # 调用保单生成逻辑
            generate_policy_docx(template_doc, data)

            # 保存为临时文件供下载
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                output_path = tmp.name
                template_doc.save(output_path)

            # 显示短信摘要（可选）
            st.success("✅ 保单生成成功！")
            st.download_button("📥 下载生成的中文保单", data=open(output_path, "rb"), file_name="中文保单.docx")
            st.subheader("📋 提取字段预览")
            st.json(data)
            
            # 可以加一个调试用文本区查看 OCR 文本
            with st.expander("🧾 OCR原始文本（调试用）"):
                st.text(ocr_text)

        
        except Exception as e:
            st.error(f"❌ 出错了：{e}")
