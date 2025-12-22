import json
import pdfplumber
import camelot


def extract_paragraph_blocks(pdf_path: str):
    blocks = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if not text:
                continue

            for line in text.split("\n"):
                line = line.strip()
                if not line:
                    continue

                blocks.append({
                    "type": "para",
                    "p": page_idx,
                    "t": line
                })

    return blocks


def extract_table_blocks(pdf_path: str):
    blocks = []

    tables = camelot.read_pdf(
        pdf_path,
        pages="all",
        flavor="lattice"  # 有格線的表格
    )

    for table in tables:
        df = table.df

        if df.shape[0] < 2:
            continue

        header = df.iloc[0].tolist()
        rows = df.iloc[1:].values.tolist()

        blocks.append({
            "type": "table",
            "p": table.page,
            "header": header,
            "rows": rows
        })

    return blocks


def build_article(pdf_path: str, topic: str):
    para_blocks = extract_paragraph_blocks(pdf_path)
    table_blocks = extract_table_blocks(pdf_path)

    blocks = para_blocks + table_blocks

    # 依 page 排序，方便人工比對 PDF
    blocks.sort(key=lambda b: (b["p"], b["type"]))

    return {
        "topic": topic,
        "blocks": blocks
    }


def main():
    #pdf_path = "input.pdf"
    pdf_path = "input/工作規則.pdf"
    topic = "第五條"

    article = build_article(pdf_path, topic)

    with open("01_articles.json", "w", encoding="utf-8") as f:
        json.dump(article, f, ensure_ascii=False, indent=2)

    print("Step1 完成：已輸出 01_articles.json")


if __name__ == "__main__":
    main()
