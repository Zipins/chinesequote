import streamlit as st
import os
import tempfile
from utils.parse_quote import extract_text_from_file, parse_quote_from_text
from utils.generate_policy import generate_policy_docx

st.set_page_config(page_title="中文保单生成器", layout="wide")
st.title("📄 中文车险保单生成系统")

uploaded_file = st.file_uploader("上传保险报价 PDF 或图片", type=["pdf", "png", "jpg", "jpeg"])

if uploaded_file:
    with st.spinner("正在识别文件内容..."):
        try:
            file_bytes = uploaded_file.read()
            filename = uploaded_file.name
            text = extract_text_from_file(file_bytes, filename)
            data = parse_quote_from_text(text)

            st.subheader("🧾 结构化字段提取结果")
            st.json(data)

            # 生成 Word 保单文件
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                output_path = tmp.name
            try:
                generate_policy_docx(data, output_path)
                with open(output_path, "rb") as f:
                    st.download_button("📥 下载中文保单 Word 文件", f, file_name="中文保单_客户名.docx")
            except Exception as e:
                st.error(f"❌ 生成失败：{e}")
            finally:
                os.unlink(output_path)
        except Exception as e:
            st.error(f"❌ 处理失败：{e}")
