# -*- coding: utf-8 -*-
from pypdf import PdfReader
import os
base = os.path.dirname(os.path.abspath(__file__))
pdf_path = os.path.join(base, "AI對話內容.pdf")
out_path = os.path.join(base, "AI對話內容_從PDF擷取.txt")
r = PdfReader(pdf_path)
parts = []
for p in r.pages:
    t = p.extract_text()
    parts.append(t if t else "")
text = "\n".join(parts)
with open(out_path, "w", encoding="utf-8") as f:
    f.write(text)
print("Pages:", len(r.pages), "Chars:", len(text))
