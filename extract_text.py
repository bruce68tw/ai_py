import pdfplumber
from utils.debug_writer import write_debug

def extract_lines(pdf_path: str):
    lines = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            text = page.extract_text(
                x_tolerance=2,
                y_tolerance=2,
                layout=True
            )
            if not text:
                continue

            for idx, raw in enumerate(text.split("\n"), start=1):
                line = raw.strip()
                if line:
                    lines.append({
                        "page": page_num,
                        "line_no": idx,
                        "text": line
                    })

    write_debug("00_raw_lines", lines)
    return lines
