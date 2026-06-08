import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"

path = ROOT / "output/DSD_Mukkamala_March_2026.pptx"
with zipfile.ZipFile(path) as z:
    root = ET.fromstring(z.read("ppt/slides/slide7.xml"))

for tag in ("sp", "graphicFrame", "cxnSp", "grpSp"):
    elems = root.findall(f".//{{{P}}}{tag}")
    print(f"{tag}: {len(elems)}")

for sp in root.findall(f".//{{{P}}}sp"):
    cnv = sp.find(f".//{{{P}}}cNvPr")
    name = cnv.get("name") if cnv is not None else "?"
    xfrm = sp.find(f".//{{{A}}}xfrm")
    if xfrm is None:
        continue
    off = xfrm.find(f"{{{A}}}off")
    ext = xfrm.find(f"{{{A}}}ext")
    if off is None or ext is None:
        continue
    x, y = int(off.get("x", 0)), int(off.get("y", 0))
    cx, cy = int(ext.get("cx", 0)), int(ext.get("cy", 0))
    geom = sp.find(f".//{{{A}}}prstGeom")
    prst = geom.get("prst") if geom is not None else ""
    # flag flat horizontal bars
    if cy < 200000 and cx > 2000000:
        print(f"H-BAR {name} y={y} cx={cx} cy={cy} prst={prst}")
    print(f"  {name}: y={y} h={cy} w={cx} prst={prst}")

for gf in root.findall(f".//{{{P}}}graphicFrame"):
    cnv = gf.find(f".//{{{P}}}cNvPr")
    name = cnv.get("name") if cnv is not None else "?"
    xfrm = gf.find(f".//{{{A}}}xfrm")
    off = xfrm.find(f"{{{A}}}off") if xfrm is not None else None
    ext = xfrm.find(f"{{{A}}}ext") if xfrm is not None else None
    if off is not None and ext is not None:
        print(f"  GF {name}: y={off.get('y')} h={ext.get('cy')} w={ext.get('cx')}")

# tblPr
for tbl in root.findall(f".//{{{A}}}tbl"):
    tbl_pr = tbl.find(f"{{{A}}}tblPr")
    if tbl_pr is not None:
        print("tblPr:", ET.tostring(tbl_pr, encoding="unicode")[:300])
