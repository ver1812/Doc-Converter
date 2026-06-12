import os
import fitz  # PyMuPDF

_PROFILES: dict[str, dict] = {
    "small":  dict(garbage=2, deflate=True),
    "medium": dict(garbage=3, deflate=True, clean=True, deflate_images=True),
    "full":   dict(garbage=4, deflate=True, clean=True, deflate_images=True, deflate_fonts=True),
}


def compress_pdf(input_path: str, output_path: str, level: str = "medium") -> dict:
    doc = fitz.open(input_path)
    doc.save(output_path, **_PROFILES[level])
    doc.close()

    orig = os.path.getsize(input_path)
    comp = os.path.getsize(output_path)
    return {
        "original_kb": orig // 1024,
        "compressed_kb": comp // 1024,
        "saved_pct": max(0.0, round((1 - comp / orig) * 100, 1)),
    }
