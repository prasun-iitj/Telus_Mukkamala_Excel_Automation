import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"

path = ROOT / "DSD_Mukkamala_February 2026_orig.pptx"
with zipfile.ZipFile(path) as z:
    root = ET.fromstring(z.read("ppt/slides/slide7.xml"))

for sp in root.findall(f".//{{{P}}}sp"):
    cnv = sp.find(f".//{{{P}}}cNvPr")
    if cnv is None or "Arrow" not in cnv.get("name", ""):
        continue
    sp_pr = sp.find(f".//{{{P}}}spPr")
    print("===", cnv.get("name"), "===")
    for child in sp_pr:
        tag = child.tag.split("}")[-1]
        if tag in ("xfrm", "prstGeom", "ln", "solidFill", "noFill"):
            print(" ", tag, child.attrib)
            if tag == "ln":
                for sub in child:
                    print("   ", sub.tag.split("}")[-1], sub.attrib)
