import streamlit as st
import os
from utils.parse_quote import extract_text_from_file, parse_quote_from_text
from utils.generate_policy import generate_policy_docx
import tempfile

st.set_page_config(page_title="中文保单生成系统")
st.title("📄 中文保单生成系统")

uploaded_file = st.file_uploader("上传报价文件（PDF、JPG、PNG）")

if uploaded_file:
    filename = uploaded_file.name
    file_bytes = uploaded_file.read()

    try:
        text = extract_text_from_file(file_bytes, filename)
        data = parse_quote_from_text(text)

        with st.expander("结构化字段提取结果"):
            st.json(data, expanded=True)

        if st.button("生成中文保单"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                output_path = tmp.name

            template_path = "template/保单范例.docx"
            generate_policy_docx(data, template_path, output_path)
            st.success("✅ 保单生成成功！")
            st.download_button("📥 下载中文保单", data=open(output_path, "rb").read(), file_name="中文保单_客户名.docx")

    except Exception as e:
        st.error(f"❌ 生成失败：{str(e)}")

