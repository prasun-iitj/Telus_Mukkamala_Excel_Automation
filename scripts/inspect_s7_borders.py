"""Compare table cell border XML on slide 7."""
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}


def cell_border_summary(path: str) -> None:
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("ppt/slides/slide7.xml"))

    print(f"=== {path} ===")
    tables = root.findall(".//a:tbl", NS)
    for ti, tbl in enumerate(tables):
        print(f"Table {ti}:")
        for ri, tr in enumerate(tbl.findall("a:tr", NS)):
            for ci, tc in enumerate(tr.findall("a:tc", NS)):
                tc_pr = tc.find("a:tcPr", NS)
                if tc_pr is None:
                    continue
                borders = {}
                for side in ("lnL", "lnR", "lnT", "lnB"):
                    ln = tc_pr.find(f"a:{side}", NS)
                    if ln is not None:
                        w = ln.get("w", "?")
                        solid = ln.find("a:solidFill/a:srgbClr", NS)
                        color = solid.get("val") if solid is not None else "?"
                        borders[side] = f"w={w} c={color}"
                if borders:
                    text = "".join(tc.itertext()).strip()[:12]
                    print(f"  r{ri}c{ci} '{text}': {borders}")

    # Rectangle 7 line props
    for sp in root.findall(".//p:sp", {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}):
        cnv = sp.find(".//p:cNvPr", {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"})
        if cnv is not None and cnv.get("name") == "Rectangle 7":
            ln = sp.find(".//a:ln", NS)
            if ln is not None:
                print("Rectangle 7 outline:", dict(ln.attrib))
                grad = ln.find("a:gradFill", NS)
                if grad is not None:
                    stops = grad.findall(".//a:gs", NS)
                    print("  gradient stops:", [(gs.get("pos"), (gs.find('a:srgbClr', NS) or {}).get('val')) for gs in stops])


for name in [
    "DSD_Mukkamala_February 2026_orig.pptx",
    "output/DSD_Mukkamala_March_2026.pptx",
]:
    p = ROOT / name
    if p.exists():
        cell_border_summary(str(p))
        print()
