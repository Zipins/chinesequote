# streamlit_policy_generator.py

import streamlit as st
import os
from utils.parse_quote import parse_quote_from_file
from utils.generate_policy import generate_policy_docx
import tempfile
import base64

st.set_page_config(page_title="\u4e2d\u6587\u4fdd\u5355\u751f\u6210\u7cfb\u7edf", layout="centered")
st.title("\ud83d\udcc4 \u4e2d\u6587\u4fdd\u5355\u751f\u6210\u7cfb\u7edf")
st.markdown("#### \u4e0a\u4f20\u4fdd\u4ef7\u6587\u4ef6\uff08PDF\u3001JPG\u3001PNG\uff09")

uploaded_file = st.file_uploader("Upload file", type=["pdf", "png", "jpg", "jpeg"])

if uploaded_file is not None:
    with tempfile.NamedTemporaryFile(delete=False, suffix="." + uploaded_file.name.split(".")[-1]) as tmp_file:
        tmp_file.write(uploaded_file.read())
        tmp_path = tmp_file.name

    try:
        with st.spinner("\u6b63\u5728\u89e3\u6790\u4fdd\u4ef7\u6587\u4ef6..."):
            parsed_data = parse_quote_from_file(tmp_path)

        st.success("\u89e3\u6790\u6210\u529f，\u6b63\u5728\u751f\u6210\u4fdd\u5355...")
        output_path = os.path.join(tempfile.gettempdir(), "\u4e2d\u6587\u4fdd\u5355.docx")
        generate_policy_docx(parsed_data, output_path)

        with open(output_path, "rb") as file:
            btn = st.download_button(
                label="\u4e0b\u8f7d\u4e2d\u6587\u4fdd\u5355",
                data=file,
                file_name="\u4e2d\u6587\u4fdd\u5355.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        st.markdown("---")
        st.markdown("#### \u77ed\u4fe1\u6458\u8981")
        sms = f"\u4f60\u597d\uff0c\u4f60\u7684\u4fdd\u9669\u4ef7\u683c\u5df2\u7ecf\u751f\u6210\uff0c\u4fdd\u9669\u516c\u53f8\uff1a{parsed_data['company']}\uff0c\u6bcf{parsed_data['policy_term']}\u4fdd\u8d39{parsed_data['total_premium']}\uff0c\u8be6\u7ec6\u4fdd\u969c\u8bf7\u89c1\u9644\u4ef6\u6587\u6863\u3002\uff08\u672c\u4ef7\u683c\u5305\u62ec{len(parsed_data.get('vehicles', []))}\u8f86\u8f66\uff09"
        st.code(sms, language="text")

    except Exception as e:
        st.error(f"\u751f\u6210\u5931\u8d25\uff1a {e}")

