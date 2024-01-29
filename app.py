
import streamlit as st
from docx import Document
import time
import mammoth
import markdownify
from io import BytesIO

def extract_outline(doc):
    # 转化Word文档为HTML
    result = mammoth.convert_to_html(doc, convert_image = mammoth.images.img_element(convert_imgs))
    # 获取HTML内容
    html = result.value
    # 转化HTML为Markdown
    md = markdownify.markdownify(html,heading_style="ATX")
    # with open("./docx_to_html.html",'w',encoding='utf-8') as html_file,open("./docx_to_md.md","w",encoding='utf-8') as md_file:
    #     html_file.write(html)
    #     md_file.write(md)
    md_result = mammoth.convert_to_markdown(doc)
    return md_result.value
            

# 转存Word文档内的图片,似乎存在问题，等待修复
def convert_imgs(image):
    with image.open() as image_bytes:
        file_suffix = image.content_type.split("/")[1]
        path_file = "./img/{}.{}".format(str(time.time()),file_suffix)
        with open(path_file, 'wb') as f:
            f.write(image_bytes.read())
    return {"src":path_file}

      
def docx2outline():
    st.title("Docx 文件大纲提取")
    uploaded_file = st.file_uploader("上传一个 .docx 文件", type="docx")
    if uploaded_file is not None:
        outline = extract_outline(uploaded_file)
        st.header("提取的大纲文本：")
        with st.expander("点击收起/展开Markdown内容", expanded=True):
            st.markdown(outline)
            st.code(outline)


if __name__ == "__main__":
    docx2outline()