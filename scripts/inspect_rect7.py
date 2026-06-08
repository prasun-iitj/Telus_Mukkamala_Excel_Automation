import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"

for fname in ["DSD_Mukkamala_February 2026_orig.pptx", "output/DSD_Mukkamala_March_2026.pptx"]:
    with zipfile.ZipFile(ROOT / fname) as z:
        root = ET.fromstring(z.read("ppt/slides/slide7.xml"))
    for sp in root.findall(f".//{{{P}}}sp"):
        cnv = sp.find(f".//{{{P}}}cNvPr")
        if cnv is None or cnv.get("name") != "Rectangle 7":
            continue
        sp_pr = sp.find(f"{{{P}}}spPr")
        print(f"=== {fname} Rectangle 7 ===")
        print(ET.tostring(sp_pr, encoding="unicode")[:2000])
