"""Inspect graphicFrame and table outer borders on slide 7."""
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"


def inspect(path: str) -> None:
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("ppt/slides/slide7.xml"))

    print(f"=== {path} ===")
    for gf in root.findall(f".//{{{P}}}graphicFrame"):
        cnv = gf.find(f".//{{{P}}}cNvPr")
        name = cnv.get("name") if cnv is not None else "?"
        xfrm = gf.find(f".//{{{A}}}xfrm")
        off = xfrm.find(f"{{{A}}}off", A) if xfrm is not None else None
        ext = xfrm.find(f"{{{A}}}ext", A) if xfrm is not None else None
        pos = (off.get("x"), off.get("y")) if off is not None else None
        size = (ext.get("cx"), ext.get("cy")) if ext is not None else None

        # outer shape properties on graphic frame
        sp_pr = gf.find(f"{{{P}}}spPr")
        ln = sp_pr.find(f"{{{A}}}ln") if sp_pr is not None else None
        ln_info = None
        if ln is not None:
            ln_info = {"w": ln.get("w")}
            solid = ln.find(f"{{{A}}}solidFill/{{{A}}}srgbClr")
            if solid is not None:
                ln_info["color"] = solid.get("val")

        tbl = gf.find(f".//{{{A}}}tbl")
        rows = tbl.findall(f"{{{A}}}tr") if tbl is not None else []
        print(f"  {name}: pos={pos} size={size} rows={len(rows)} frame_ln={ln_info}")

        if tbl is not None and rows:
            # last row bottom borders
            last = rows[-1]
            for ci, tc in enumerate(last.findall(f"{{{A}}}tc")):
                tc_pr = tc.find(f"{{{A}}}tcPr")
                if tc_pr is None:
                    continue
                ln_b = tc_pr.find(f"{{{A}}}lnB")
                if ln_b is not None:
                    text = "".join(tc.itertext()).strip()[:10]
                    print(f"    last row c{ci} '{text}' lnB w={ln_b.get('w')}")

            # first row top borders (table 1 title)
            first = rows[0]
            for ci, tc in enumerate(first.findall(f"{{{A}}}tc")):
                tc_pr = tc.find(f"{{{A}}}tcPr")
                if tc_pr is None:
                    continue
                ln_t = tc_pr.find(f"{{{A}}}lnT")
                if ln_t is not None:
                    text = "".join(tc.itertext()).strip()[:20]
                    print(f"    first row c{ci} '{text}' lnT w={ln_t.get('w')}")


for name in [
    "DSD_Mukkamala_February 2026_orig.pptx",
    "output/DSD_Mukkamala_March_2026.pptx",
]:
    p = ROOT / name
    if p.exists():
        inspect(str(p))
        print()
