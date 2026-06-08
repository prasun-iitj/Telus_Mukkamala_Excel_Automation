"""Structural diff of slide 7 XML (ignore text)."""
import re
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def normalize(xml: str) -> str:
    xml = re.sub(r">[^<]+<", "><", xml)  # strip text
    xml = re.sub(r"\s+", "", xml)
    return xml


def size(path: str) -> int:
    with zipfile.ZipFile(path) as z:
        return len(z.read("ppt/slides/slide7.xml"))


def main() -> None:
    paths = {
        "orig": ROOT / "DSD_Mukkamala_February 2026_orig.pptx",
        "march": ROOT / "output/DSD_Mukkamala_March_2026.pptx",
        "feb": ROOT / "output/DSD_Mukkamala_February_2026.pptx",
    }
    for k, p in paths.items():
        print(k, "slide7 bytes:", size(str(p)))

    with zipfile.ZipFile(paths["orig"]) as z:
        orig = normalize(z.read("ppt/slides/slide7.xml").decode())
    with zipfile.ZipFile(paths["march"]) as z:
        march = normalize(z.read("ppt/slides/slide7.xml").decode())

    # count tags
    for tag in ["p:sp", "a:tbl", "a:tr", "a:tc", "upArrow", "p:graphicFrame"]:
        print(f"orig {tag}:", orig.count(tag), "march:", march.count(tag))


if __name__ == "__main__":
    main()
