import streamlit as st
from bs4 import BeautifulSoup
from pptx import Presentation
from pptx.util import Inches, Pt
import io
import time
import os
from playwright.sync_api import sync_playwright

@st.cache_resource
def install_playwright():
    os.system("playwright install chromium")

install_playwright()



# --- 🚀 华丽的 Web 界面构建 ---
st.set_page_config(page_title="HTML to PPT 转换器", page_icon="🚀", layout="wide")

with st.sidebar:
    st.image("https://img.icons8.com/color/96/000000/microsoft-powerpoint-2019--v1.png", width=60)
    st.header("⚙️ 转换器引擎状态")
    st.success("✅ 视觉抓拍引擎 (Playwright) 已挂载")
    st.success("✅ DOM无限穿透引擎 已就绪")
    st.markdown("---")
    st.markdown("### 💡 智能识别能力：")
    st.markdown("- **无限穿透**：无视前端外壳嵌套，直击数据底层\n- **文本结构**：各级标题、段落排版\n- **复杂表格**：原生转换为 PPT 表格格式\n- **动态图表**：自动定位、滚动并高清抓拍")
    st.markdown("---")
    st.caption("运行环境：Python 3.12 虚拟隔离舱")

st.title("🚀 HTML to PPT 转换器 (智能容器识别版)")
st.markdown("欢迎使用全自动汇报生成流水线。只需拖入 HTML 文件，系统将自动穿透 DOM 节点，完成动态图表渲染与幻灯片排版。")
st.markdown("---")

uploaded_file = st.file_uploader("📂 请上传需要转换的HTML文件：", type=["html", "htm"])

if uploaded_file is not None:
    original_filename = uploaded_file.name
    output_filename = original_filename.replace(".html", ".pptx").replace(".htm", ".pptx")
    
    html_bytes = uploaded_file.read()
    html_string = html_bytes.decode("utf-8")
    
    st.info(f"📄 文件 `{original_filename}` 载入成功！转换流水线已待命。")
    
    if st.button("⚡ 启动引擎并生成 PPT", type="primary"):
        with st.spinner('底层引擎全速运转中！正在无限穿透嵌套层级，请稍候...'):
            pptx_file = html_to_pptx_with_charts(html_string)
            
        st.balloons()
        st.success(f"🎉 转换圆满完成！已生成：**{output_filename}**")
        
        st.download_button(
            label="⬇️ 点击下载 PPTX 文件",
            data=pptx_file,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
