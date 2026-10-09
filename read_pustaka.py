from docx import Document

doc = Document("LAPORAN_PW_FAIRUZ_TEMP_v2.docx")

found = False
for i, p in enumerate(doc.paragraphs):
    if "DAFTAR PUSTAKA" in p.text and p.style.name.startswith("Heading"):
        found = True
        print(f"--- DAFTAR PUSTAKA START (Para {i}) ---")
        continue
    
    if found:
        # Stop at the next heading
        if p.style.name.startswith("Heading") and p.text.strip():
            break
        
        text = p.text.strip()
        if text:
            print(f"[{i}] {text}")
