import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt
import re
import io

# Funções de formatação (as mesmas que você já validou)
def set_rtl(paragraph):
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    bidi = pPr.find(qn('w:bidi'))
    if bidi is None:
        bidi = OxmlElement('w:bidi')
        bidi.set(qn('w:val'), '1')
        pPr.append(bidi)
    else:
        bidi.set(qn('w:val'), '1')
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT

def process_docx(input_file):
    doc = Document(input_file)
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text: continue
        
        # Header Detection
        try:
            if re.match(r'^\d+\.\s', text):
                para.style = 'Heading 1'
            elif re.match(r'^[A-Z\s]{5,}$', text):
                para.style = 'Heading 2'
        except:
            pass
        
        # Body Text
        if para.style.name not in ['Heading 1', 'Heading 2']:
            para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            for run in para.runs:
                run.font.name = 'Arial'
                run.font.size = Pt(11)

        # RTL Handling
        if re.search(r'[\u0600-\u06FF]', text):
            set_rtl(para)
            
    # Salva o arquivo em um objeto de memória (BytesIO) para o download
    target_stream = io.BytesIO()
    doc.save(target_stream)
    target_stream.seek(0)
    return target_stream

# Interface Streamlit
st.set_page_config(page_title="Docx RTL Formatter", page_icon="📄")
st.title("📄 Professional Docx Formatter")
st.markdown("Automate your document styling and Arabic (RTL) alignment in seconds.")

uploaded_file = st.file_uploader("Upload your input.docx file", type=["docx"])

if uploaded_file is not None:
    if st.button("🚀 Format Document"):
        with st.spinner('Processing...'):
            # Processa o arquivo enviado
            processed_file = process_docx(uploaded_file)
            
            st.success("Success! Your document is ready.")
            
            # Botão de Download
            st.download_button(
                label="📥 Download Formatted File",
                data=processed_file,
                file_name="formatted_output.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )