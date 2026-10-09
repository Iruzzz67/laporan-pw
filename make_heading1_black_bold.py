from docx import Document
from docx.shared import RGBColor
from docx.oxml.ns import qn
import sys

def set_black_color_el(font_element):
    if font_element is None: return
    color_el = font_element.find(qn('w:color'))
    if color_el is not None:
        color_el.set(qn('w:val'), '000000')
        for attr in ['w:themeColor', 'w:themeTint', 'w:themeShade']:
            if qn(attr) in color_el.attrib:
                del color_el.attrib[qn(attr)]
    else:
        from docx.oxml.shared import OxmlElement
        color_el = OxmlElement('w:color')
        color_el.set(qn('w:val'), '000000')
        font_element.append(color_el)

def set_bold_el(font_element):
    if font_element is None: return
    b_el = font_element.find(qn('w:b'))
    if b_el is not None:
        b_el.set(qn('w:val'), '1')
    else:
        from docx.oxml.shared import OxmlElement
        b_el = OxmlElement('w:b')
        font_element.append(b_el)
        
    bCs_el = font_element.find(qn('w:bCs'))
    if bCs_el is not None:
        bCs_el.set(qn('w:val'), '1')
    else:
        from docx.oxml.shared import OxmlElement
        bCs_el = OxmlElement('w:bCs')
        font_element.append(bCs_el)

def process(doc_path, save_path):
    try:
        doc = Document(doc_path)
        
        # 1. Change Heading 1 Style
        if 'Heading 1' in doc.styles:
            style = doc.styles['Heading 1']
            if style.font._element is not None:
                set_black_color_el(style.font._element)
                set_bold_el(style.font._element)
                
        # 2. Change all runs in Heading 1 paragraphs
        count = 0
        for p in doc.paragraphs:
            if p.style.name == 'Heading 1':
                for run in p.runs:
                    if run.font._element is not None:
                        set_black_color_el(run.font._element)
                        set_bold_el(run.font._element)
                
                # Update paragraph mark as well
                pPr = p._element.get_or_add_pPr()
                rPr = pPr.find(qn('w:rPr'))
                if rPr is None:
                    from docx.oxml.shared import OxmlElement
                    rPr = OxmlElement('w:rPr')
                    pPr.append(rPr)
                set_black_color_el(rPr)
                set_bold_el(rPr)
                
                count += 1
                
        doc.save(save_path)
        print(f"Processed {count} Heading 1 paragraphs in {save_path}")
    except Exception as e:
        print(f"Error: {e}")

process("LAPORAN_PW_FAIRUZ_TEMP_v4.docx", "LAPORAN_PW_FAIRUZ_TEMP_v4.docx")
process("LAPORAN_PW_FAIRUZ_REVISI_20 Mei (2) (1)_REVISED.docx", "LAPORAN_PW_FAIRUZ_REVISI_20 Mei (2) (1)_REVISED.docx")
