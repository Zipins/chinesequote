import logging
logging.basicConfig(level=logging.INFO)
import streamlit as st
import os
from utils.parse_quote import extract_quote_data
from utils.generate_policy import generate_policy_docx
from tempfile import NamedTemporaryFile
from docx import Document
import traceback

st.set_page_config(page_title="中文保单生成器", layout="centered")
st.title("📄 中文保单生成器")
st.markdown("请上传保险 quote 文件（支持 PDF、扫描件或图片），系统将自动识别内容并生成中文保单 Word 文档。")

uploaded_file = st.file_uploader("上传 PDF 或 图片文件", type=["pdf", "png", "jpg", "jpeg"])
custom_name = st.text_input("输出文件名称（可选）：", value="中文保单")
show_ocr = st.checkbox("🔍 显示原始识别文本（用于调试）")

if uploaded_file and st.button("生成保单"):
    try:
        with st.spinner("正在识别内容并生成保单，请稍候..."):
            st.subheader("📘 调试日志")
            st.write(f"📄 文件名：{uploaded_file.name}")
            st.write("🔍 调用 extract_quote_data() 开始...")

            # 调用后端函数
            data, ocr_text = extract_quote_data(uploaded_file, return_raw_text=True)

            if show_ocr:
                st.subheader("🧾 原始 OCR 文本（仅供调试）")
                st.code(ocr_text[:10000], language="text")

            template_path = "template/保单范例.docx"
            if not os.path.exists(template_path):
                st.error("❌ 保单模板文件不存在，请检查 template/ 文件夹。")
                st.stop()

            doc = Document(template_path)
            generate_policy_docx(doc, data)

            with NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                doc.save(tmp.name)
                tmp_path = tmp.name

            st.success("✅ 保单生成成功！")
            with open(tmp_path, "rb") as f:
                st.download_button("📥 下载保单 Word 文件", f, file_name=f"{custom_name}.docx")

    except Exception as e:
        st.error(f"❌ 出错了：{e}")
        st.code(traceback.format_exc(), language="python")
