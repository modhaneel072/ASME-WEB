"""One-off utility to extract embedded headshot images from Roles ASME.pdf
into static/images/executive/, named per page so a human can rename them
to match what's referenced in MANUAL_EXECUTIVE_PROFILES."""

import fitz
from pathlib import Path

PDF_PATH = Path(r"C:\Users\modha\Downloads\Roles ASME.pdf")
OUT_DIR = Path(__file__).resolve().parent.parent / "static" / "images" / "executive"
OUT_DIR.mkdir(parents=True, exist_ok=True)

doc = fitz.open(PDF_PATH)
for page_num, page in enumerate(doc, start=1):
    images = page.get_images(full=True)
    for img_idx, img in enumerate(images, start=1):
        xref = img[0]
        base = doc.extract_image(xref)
        ext = base["ext"]
        data = base["image"]
        out_path = OUT_DIR / f"_pdf_p{page_num}_i{img_idx}.{ext}"
        out_path.write_bytes(data)
        print(f"page {page_num} image {img_idx}: {out_path.name} ({len(data)} bytes, {base.get('width')}x{base.get('height')})")
