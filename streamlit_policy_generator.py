import streamlit as st
import tempfile
import os
from utils.parse_quote import extract_text_from_file, parse_quote_from_text
from utils.generate_policy import generate_policy_docx

def main():
    st.set_page_config(page_title="中文保险保单生成器", layout="centered")
    st.title("📄 中文保险保单生成器")

    uploaded_file = st.file_uploader("请上传保险报价单（支持 PDF、JPG、PNG）", type=["pdf", "png", "jpg", "jpeg"])

    output_filename = st.text_input("生成的保单文件名（可选，不含 .docx 后缀）", value="中文保单_客户名")

    if uploaded_file:
        st.success(f"文件 {uploaded_file.name} 上传成功，正在解析...")

        try:
            # 保存上传文件到临时路径
            with tempfile.NamedTemporaryFile(delete=False) as tmp_file:
                tmp_file.write(uploaded_file.read())
                tmp_file_path = tmp_file.name

            # 提取文本并结构化
            raw_text = extract_text_from_file(tmp_file_path, uploaded_file.name)
            data = parse_quote_from_text(raw_text)

            st.subheader("✅ 结构化字段提取结果")
            st.json(data)

            # 生成保单 Word 文件
            output_path = f"{output_filename.strip()}.docx"
            generate_policy_docx(data, output_path)

            with open(output_path, "rb") as f:
                st.download_button(
                    label="📥 下载生成的中文保单",
                    data=f,
                    file_name=output_path,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

            # 可选短信模板
            st.subheader("📱 可发送给客户的短信内容")
            sms = f"您好，这是我们根据您提供的信息生成的保险报价概览，详情请见附件保单文件《{output_path}》。如有问题欢迎随时联系！"
            st.code(sms)

        except Exception as e:
            st.error(f"❌ 生成失败：{str(e)}")
        finally:
            if os.path.exists(tmp_file_path):
                os.remove(tmp_file_path)

if __name__ == "__main__":
    main()
