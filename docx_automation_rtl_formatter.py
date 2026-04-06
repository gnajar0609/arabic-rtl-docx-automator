from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt
import re
import os

def set_rtl(paragraph):
    """
    Applies RTL (Right-to-Left) and Bidi flags directly to the OOXML 
    to support Arabic and other bidirectional languages.
    """
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    
    # Check if bidi element already exists
    bidi = pPr.find(qn('w:bidi'))
    if bidi is None:
        # Create and append the bidi element manually using OxmlElement
        bidi = OxmlElement('w:bidi')
        bidi.set(qn('w:val'), '1')
        pPr.append(bidi)
    else:
        bidi.set(qn('w:val'), '1')
    
    # Align the paragraph to the right
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT

def format_document(input_path, output_path):
    """
    Main engine to process the Word document:
    1. Detects headers (Heading 1/2) using Regex.
    2. Formats body text to Arial 11pt, Justified.
    3. Automatically triggers RTL for Arabic text detection.
    """
    if not os.path.exists(input_path):
        print(f"❌ ERROR: File '{input_path}' not found!")
        return

    doc = Document(input_path)
    
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text: 
            continue

        # 1. Header Detection
        try:
            # Matches "1. Title"
            if re.match(r'^\d+\.\s', text):
                para.style = 'Heading 1'
            # Matches "ALL CAPS TITLES" (at least 5 chars)
            elif re.match(r'^[A-Z\s]{5,}$', text):
                para.style = 'Heading 2'
        except Exception:
            # Fallback if specific styles are missing in the template
            pass
        
        # 2. Body Text Formatting
        # Only apply if it's not a detected header
        if para.style.name not in ['Heading 1', 'Heading 2']:
            para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            for run in para.runs:
                run.font.name = 'Arial'
                run.font.size = Pt(11)

        # 3. RTL Handling (Arabic Unicode Range: \u0600-\u06FF)
        if re.search(r'[\u0600-\u06FF]', text):
            set_rtl(para)

    doc.save(output_path)
    print(f"✅ Success! File generated: {output_path}")

if __name__ == "__main__":
    # Ensure 'input.docx' exists in the same directory
    format_document('input.docx', 'formatted_output.docx')