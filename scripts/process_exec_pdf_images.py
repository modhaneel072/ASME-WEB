"""Convert the raw PDF-extracted images into final headshot files.

Three of the extracted images are clean standalone headshots — those just
get renamed/converted to JPG. Two are Discord/Teams screenshots with the
photo embedded in the bottom-left; we crop the photo region out.
"""

from pathlib import Path
from PIL import Image

EXEC_DIR = Path(__file__).resolve().parent.parent / "static" / "images" / "executive"

# (source filename, output filename, crop box or None for full image, JPG quality)
JOBS = [
    # Clean standalone photos — no crop, just convert to JPG
    ("_pdf_p1_i1.jpeg", "henry_strauss.jpg",      None,                          92),
    ("_pdf_p1_i2.jpeg", "dawson_fish.jpg",        None,                          92),
    ("_pdf_p2_i1.jpeg", "nathan_fish.jpg",        None,                          92),
    # Screenshots — crop to just the headshot in the bottom-left of each
    # (boxes determined from the 511x462 / 502x492 PNGs in the PDF)
    ("_pdf_p4_i1.png",  "adelai_kaiser.jpg",      (20, 232, 275, 460),           90),
    ("_pdf_p5_i1.png",  "heriberto_salgado.jpg",  (15, 240, 240, 488),           90),
]

for src_name, out_name, box, quality in JOBS:
    src = EXEC_DIR / src_name
    if not src.exists():
        print(f"!! missing {src_name}")
        continue
    img = Image.open(src).convert("RGB")
    if box is not None:
        img = img.crop(box)
    out = EXEC_DIR / out_name
    img.save(out, "JPEG", quality=quality, optimize=True)
    print(f"-> {out.name} ({img.width}x{img.height})")

# Clean up: remove all _pdf_* extraction artifacts
for f in EXEC_DIR.glob("_pdf_*"):
    f.unlink()
    print(f"removed {f.name}")
