from docx import Document

doc = Document("LAPORAN_PW_FAIRUZ_REVISI_20 Mei (2) (1)_REVISED.docx")
found = False
for p in doc.paragraphs:
    if "DAFTAR PUSTAKA" in p.text and p.style.name.startswith("Heading"):
        found = True
        continue
    
    if found:
        if p.style.name.startswith("Heading") and p.text.strip():
            break
        text = p.text.strip()
        if text and text.startswith("["):
            print("Before:")
            print(p.text)
            # Remove the hyperlink element and the "(Tautan: " ")" runs
            for child in list(p._p):
                if child.tag.endswith('hyperlink'):
                    p._p.remove(child)
            
            # Now remove the "(Tautan: " and ")" from text
            # Easiest way without losing bold/italic is to replace the text in runs
            for r in p.runs:
                if "(Tautan: " in r.text:
                    r.text = r.text.replace(" (Tautan: ", "")
                if ")" in r.text and len(r.text.strip()) == 1:
                    r.text = r.text.replace(")", "")
            print("After:")
            print(p.text)
