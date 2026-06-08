"""Inspect MoM column and arrow shape XML on slide 7."""
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def dump(path: str) -> None:
    with zipfile.ZipFile(path) as z:
        xml = z.read("ppt/slides/slide7.xml")

    root = ET.fromstring(xml)
    print(f"=== {path} ===")

    # Arrow shapes detail
    for sp in root.findall(f".//{{{P}}}sp"):
        cnv = sp.find(f".//{{{P}}}cNvPr")
        if cnv is None:
            continue
        name = cnv.get("name", "")
        if "Arrow" in name:
            sp_pr = sp.find(f".//{{{P}}}spPr")
            xfrm = sp_pr.find(f"{{{A}}}xfrm") if sp_pr is not None else None
            geom = sp_pr.find(f"{{{A}}}prstGeom") if sp_pr is not None else None
            ln = sp_pr.find(f"{{{A}}}ln") if sp_pr is not None else None
            print(f"  ARROW {name}: prst={geom.get('prst') if geom is not None else None}")
            if xfrm is not None:
                off = xfrm.find(f"{{{A}}}off")
                ext = xfrm.find(f"{{{A}}}ext")
                print(f"    pos=({off.get('x')},{off.get('y')}) size=({ext.get('cx')},{ext.get('cy')})")
            if ln is not None:
                print(f"    ln w={ln.get('w')}")
                solid = ln.find(f"{{{A}}}solidFill/{{{A}}}srgbClr")
                if solid is not None:
                    print(f"    ln color={solid.get('val')}")

    # MoM cells - any drawing inside?
    tables = root.findall(f".//{{{A}}}tbl")
    if tables:
        tbl = tables[0]
        rows = tbl.findall(f"{{{A}}}tr")
        for ri in (2, 3, 4):
            if ri >= len(rows):
                continue
            cells = rows[ri].findall(f"{{{A}}}tc")
            if len(cells) < 5:
                continue
            tc = cells[4]
            text = "".join(tc.itertext()).strip()
            tx_body = tc.find(f"{{{A}}}txBody")
            extra = [child.tag.split("}")[-1] for child in tc if child.tag.split("}")[-1] != "txBody"]
            print(f"  MoM r{ri}c4 text='{text}' extra_children={extra}")


for name in [
    "DSD_Mukkamala_February 2026_orig.pptx",
    "output/DSD_Mukkamala_March_2026.pptx",
]:
    p = ROOT / name
    if p.exists():
        dump(str(p))
        print()
