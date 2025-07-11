import streamlit as st
from utils.parse_quote import extract_quote_data
from utils.generate_policy import generate_policy_doc
import os
import datetime
import tempfile

st.set_page_config(page_title="中文车险保单生成器", layout="centered")
st.title("📄 中文车险保单生成器")
st.markdown("---")

uploaded_file = st.file_uploader("请上传保险报价单 (PDF 或 图片)", type=["pdf", "jpg", "jpeg", "png"])

# 可选项控制
show_ocr = st.checkbox("显示 OCR 原始文本")

if uploaded_file:
    try:
        st.markdown(f"📄 文件名：`{uploaded_file.name}`")
        st.markdown("🔍 调用 `extract_quote_data()` 开始...")

        data, ocr_text = extract_quote_data(uploaded_file, return_raw_text=True)

        st.success("✅ 数据提取成功")

        with st.expander("📑 提取字段结构"):
            st.json(data)

        if show_ocr:
            with st.expander("📄 原始 OCR 文本"):
                st.text(ocr_text)

        # 输出文件名输入框
        default_name = "中文保单_客户名.docx"
        output_filename = st.text_input("输出文件名称：", value=default_name)

        if st.button("📄 生成中文保单"):
            with tempfile.TemporaryDirectory() as tmpdir:
                output_path = os.path.join(tmpdir, output_filename)
                generate_policy_doc(data, output_path)
                with open(output_path, "rb") as f:
                    st.download_button("📥 下载生成的中文保单", f, file_name=output_filename)

    except Exception as e:
        st.error(f"❌ 出错了：{str(e)}")
