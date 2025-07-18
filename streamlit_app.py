import streamlit as st
import tempfile
import os
from docx import Document
from utils.parse_quote import parse_quote_from_file
from utils.generate_policy import generate_policy_docx

st.set_page_config(page_title="中文保单生成器", layout="wide")

st.title("📄 中文车险保单生成器")

uploaded_file = st.file_uploader("请上传保险报价单（支持 PDF 或图片）", type=["pdf", "jpg", "jpeg", "png"])

if uploaded_file:
    try:
        with st.spinner("正在识别保单内容，请稍候..."):
            # 临时保存上传文件
            with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1]) as tmp:
                tmp.write(uploaded_file.read())
                tmp_path = tmp.name

            # 提取 quote 信息
            data, full_text = parse_quote_from_file(tmp_path)

            # 展示 OCR 提取内容（调试用）
            with st.expander("📑 提取字段结构（结构化数据）", expanded=False):
                st.code(data, language="json")
            with st.expander("📄 原始 OCR 文本", expanded=False):
                st.text(full_text[:8000])  # 限制最大展示长度

            # 用户自定义输出文件名
            default_filename = "中文保单_客户名.docx"
            filename = st.text_input("输出文件名", value=default_filename)

            # 生成保单按钮
            if st.button("📃 生成中文保单"):
                with st.spinner("正在生成 Word 保单..."):
                    doc = Document("template/保单范例.docx")
                    generate_policy_docx(doc, data)

                    # 保存为临时文件供下载
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as out_file:
                        doc.save(out_file.name)
                        st.success("✅ 保单生成成功！")
                        with open(out_file.name, "rb") as f:
                            st.download_button("📥 下载保单", data=f, file_name=filename, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    except Exception as e:
        st.error(f"❌ 出错了：{e}")
