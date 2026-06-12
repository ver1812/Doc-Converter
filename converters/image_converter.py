from PIL import Image
import pillow_heif

pillow_heif.register_heif_opener()

HEIC_EXTS = {".heic", ".heif"}
FORMAT_MAP = {
    "JPEG": "JPEG",
    "PNG": "PNG",
    "HEIC": "HEIF",
}


def convert_image(input_path: str, output_path: str, target_format: str) -> bool:
    fmt = FORMAT_MAP[target_format.upper()]
    img = Image.open(input_path)

    if fmt == "JPEG":
        if img.mode in ("RGBA", "P", "LA"):
            img = img.convert("RGB")
        elif img.mode != "RGB":
            img = img.convert("RGB")
        img.save(output_path, format="JPEG", quality=95)
    elif fmt == "HEIF":
        if img.mode not in ("RGB", "RGBA"):
            img = img.convert("RGB")
        heif_img = pillow_heif.from_pillow(img)
        heif_img.save(output_path)
    else:
        img.save(output_path, format=fmt)

    return True
