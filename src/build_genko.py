"""読み上げ原稿（docs/読み上げ原稿_配布用短縮版.md）を、印刷用の Word と PDF にする。
使い方：python3 src/build_genko.py  （pandoc・python-docx・LibreOffice が必要）"""
import subprocess, docx
from docx.shared import Pt, Cm, RGBColor
from docx.oxml.ns import qn

SRC = "docs/読み上げ原稿_配布用短縮版.md"
OUT = "output/読み上げ原稿_配布用短縮版.docx"
F = "BIZ UDPゴシック"

subprocess.run(["pandoc", SRC, "-o", OUT], check=True)
d = docx.Document(OUT)
for sec in d.sections:
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = Cm(2)
for st in d.styles:
    try:
        f = st.font
    except Exception:
        continue
    if f is None:
        continue
    f.name = F
    rpr = st.element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rf)
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(a), F)
    for a in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rf.get(qn(a)) is not None:
            del rf.attrib[qn(a)]
for n, s in {"Normal": 14, "Body Text": 14, "First Paragraph": 14, "Compact": 14, "Heading 1": 20, "Heading 2": 16}.items():
    try:
        d.styles[n].font.size = Pt(s)
    except KeyError:
        pass
for n in ("Heading 1", "Heading 2"):
    try:
        d.styles[n].font.color.rgb = RGBColor(0x1D, 0x46, 0x36)
    except KeyError:
        pass
for para in d.paragraphs:
    para.paragraph_format.line_spacing = 1.4
d.save(OUT)
subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", "output", OUT], check=True, capture_output=True)
print("ok")
