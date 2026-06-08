"""Inspect slide 7 XML for distortion sources."""
import re
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def inspect(path: str) -> None:
    with zipfile.ZipFile(path) as z:
        xml = z.read("ppt/slides/slide7.xml").decode("utf-8")

    print(f"=== {path} ===")
    print("shapes:", len(re.findall(r"<p:sp\b", xml)))
    print("Arrow mentions:", xml.count("Arrow"))
    print("upArrow geom:", xml.lower().count("uparrow"))
    widths = sorted({int(m) for m in re.findall(r'<a:ln w="(\d+)"', xml)})
    print("line widths (EMU):", widths[:20], "..." if len(widths) > 20 else "")
    print("tbl count:", xml.count("<a:tbl"))

    # MoM column cells in table 0 - look for drawings inside tc
    for m in re.finditer(r"<a:tc>.*?</a:tc>", xml, re.DOTALL):
        block = m.group()
        if "MoM" in block or "upArrow" in block.lower():
            print("MoM-ish cell block len:", len(block))


for name in [
    "DSD_Mukkamala_February 2026_orig.pptx",
    "output/DSD_Mukkamala_March_2026.pptx",
]:
    p = ROOT / name
    if p.exists():
        inspect(str(p))
        print()
