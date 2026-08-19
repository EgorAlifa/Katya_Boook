"""
Extracts authoritative per-image layout (display size in inches, rotation in
degrees, flip flags) straight from the original docx's word/document.xml —
pandoc's markdown loses all of this (it only gives pixel dimensions of the
raw file, not the size/rotation Word actually displays it at).

Walks every <w:drawing> in document order, resolves its r:embed relationship
id to a media filename via document.xml.rels, and records:
  - w_in / h_in   : wp:extent cx/cy (EMU -> inches) = the size Word renders
                    the picture at, which can differ from the raw file's
                    native pixel size.
  - rotation      : a:xfrm/@rot (60000ths of a degree -> degrees)
  - flipH / flipV : a:xfrm/@flipH / @flipV

Since a handful of images are reused (same file, multiple placements), the
result maps filename -> LIST of occurrences in document order; classify_book
matches image blocks to this list positionally (each source file's images
are consumed in the same left-to-right, top-to-bottom order pandoc emits
them in).
"""
import json
import re
from collections import defaultdict
from lxml import etree

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

EMU_PER_INCH = 914400


def load_rels(path):
    rels_xml = open(path, encoding="utf-8").read()
    return dict(re.findall(r'Id="(rId\d+)"[^>]*Target="media/([^"]+)"', rels_xml))


def main():
    doc = etree.parse("original_unzip/word/document.xml")
    rid2file = load_rels("original_unzip/word/_rels/document.xml.rels")

    occurrences = defaultdict(list)
    n_drawings = 0
    n_matched = 0

    for drawing in doc.iter("{%s}drawing" % NS["w"]):
        n_drawings += 1
        extent = drawing.find(".//wp:extent", NS)
        blip = drawing.find(".//a:blip", NS)
        if extent is None or blip is None:
            continue
        rid = blip.get("{%s}embed" % NS["r"])
        fname = rid2file.get(rid)
        if not fname:
            continue
        n_matched += 1
        w_in = int(extent.get("cx")) / EMU_PER_INCH
        h_in = int(extent.get("cy")) / EMU_PER_INCH

        xfrm = drawing.find(".//pic:spPr/a:xfrm", NS)
        rotation = 0.0
        flip_h = False
        flip_v = False
        if xfrm is not None:
            rot = xfrm.get("rot")
            if rot:
                rotation = int(rot) / 60000.0
            flip_h = xfrm.get("flipH") == "1"
            flip_v = xfrm.get("flipV") == "1"

        occurrences[fname].append({
            "w_in": round(w_in, 4),
            "h_in": round(h_in, 4),
            "rotation": rotation,
            "flipH": flip_h,
            "flipV": flip_v,
        })

    print(f"drawings seen: {n_drawings}, matched to media: {n_matched}, unique files: {len(occurrences)}")
    rotated = {f: occs for f, occs in occurrences.items() if any(o["rotation"] for o in occs)}
    print(f"files with a rotated occurrence: {rotated}")

    with open("image_layout.json", "w", encoding="utf-8") as f:
        json.dump(occurrences, f, ensure_ascii=False, indent=2)
    print("wrote image_layout.json")


if __name__ == "__main__":
    main()
