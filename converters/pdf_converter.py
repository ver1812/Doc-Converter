from pdf2docx import Converter


def convert_pdf_to_docx(input_path: str, output_path: str) -> bool:
    cv = Converter(input_path)
    try:
        cv.convert(output_path, start=0, end=None)
    finally:
        cv.close()
    return True
