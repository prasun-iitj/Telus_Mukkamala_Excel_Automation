"""Compare cell text XML structure after updates."""
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def para_counts(path: str) -> None:
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("ppt/slides/slide7.xml"))

    print(f"=== {path} ===")
    for ti, tbl in enumerate(root.findall(f".//{{{A}}}tbl")):
        for ri, tr in enumerate(tbl.findall(f"{{{A}}}tr")):
            for ci, tc in enumerate(tr.findall(f"{{{A}}}tc")):
                tx = tc.find(f"{{{A}}}txBody")
                if tx is None:
                    continue
                paras = tx.findall(f"{{{A}}}p")
                runs = sum(len(p.findall(f"{{{A}}}r")) for p in paras)
                text = "".join(tc.itertext()).strip()[:20]
                if runs > 1 or len(paras) > 1:
                    print(f"  t{ti} r{ri}c{ci} '{text}' paras={len(paras)} runs={runs}")


for name in [
    "DSD_Mukkamala_February 2026_orig.pptx",
    "output/DSD_Mukkamala_March_2026.pptx",
]:
    p = ROOT / name
    if p.exists():
        para_counts(str(p))
        print()
