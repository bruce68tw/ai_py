import os
import re
import json
from docx import Document

# ===== 正則定義 =====
RE_CHAPTER = re.compile(r'^第\s*([一二三四五六七八九十0-9]+)\s*章\s*(.*)')
RE_ARTICLE = re.compile(r'^第\s*([一二三四五六七八九十0-9]+)\s*條\s*(.*)')
RE_SUB = re.compile(r'^([a-z])[\.\、]?\s*(.*)')
RE_ITEM = re.compile(r'^([一二三四五六七八九十]+)[、\.]\s*(.*)')

def extract_page(paragraph):
    """
    python-docx 無法取得實際頁碼
    保留欄位，回傳 None 是正確行為
    """
    return None

def clean_body(text: str) -> str:
    """
    - 條文內部不得有換行
    - 條文之間只保留一個換行
    """
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    return "\n".join(lines)

def word_to_json(word_path, output_path):
    doc = Document(word_path)

    current_chapter_title = None   # 第0階
    current_article = None         # 第1階
    current_sub = None             # 第2階
    current_item = None            # 一、二、三 條文緩衝

    results = []

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        # ===== 第0階：章 =====
        m = RE_CHAPTER.match(text)
        if m:
            title = m.group(2).strip()
            current_chapter_title = None if title == "(刪除)" else title
            current_article = None
            current_sub = None
            current_item = None
            continue

        # ===== 第1階：條 =====
        m = RE_ARTICLE.match(text)
        if m:
            title = m.group(2).strip()
            if title == "(刪除)":
                current_article = None
                current_sub = None
                current_item = None
                continue

            full_title = title
            if current_chapter_title and current_chapter_title not in title:
                full_title = f"{current_chapter_title} {title}"

            current_article = {
                "article0": f"第{m.group(1)}條 {full_title}",
                "article": full_title,
                "page": extract_page(para),
                "body": ""
            }
            results.append(current_article)
            current_sub = None
            current_item = None
            continue

        # ===== 第2階：a,b,c =====
        m = RE_SUB.match(text)
        if m and current_article:
            title = m.group(2).strip()
            if title == "(刪除)":
                current_sub = None
                current_item = None
                continue

            full_title = title
            if current_article["article"] not in title:
                full_title = f"{current_article['article']} {title}"

            current_sub = {
                "article0": f"{m.group(1)} {full_title}",
                "article": full_title,
                "page": extract_page(para),
                "body": ""
            }
            results.append(current_sub)
            current_item = None
            continue

        # ===== 條文內容處理 =====
        target = current_sub if current_sub else current_article
        if not target:
            continue

        # 是否為「一、二、三…」條文起始
        m_item = RE_ITEM.match(text)
        if m_item:
            current_item = text.strip()
            target["body"] += current_item + "\n"
        else:
            # 條文續行 → 合併到同一條
            if current_item:
                target["body"] = target["body"].rstrip("\n")
                target["body"] += " " + text.strip() + "\n"
            else:
                # 非條列的一般說明文字
                target["body"] += text.strip() + "\n"

    # ===== 最終清洗 body =====
    for r in results:
        r["body"] = clean_body(r["body"])

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

def batch_convert(input_path, output_dir):
    if not input_path.lower().endswith(".docx"):
        raise ValueError("只接受單一 .docx 檔案")

    if not os.path.isfile(input_path):
        raise FileNotFoundError(f"找不到檔案：{input_path}")

    os.makedirs(output_dir, exist_ok=True)

    fname = os.path.basename(input_path)
    output_path = os.path.join(
        output_dir,
        fname.replace(".docx", ".json")
    )

    word_to_json(input_path, output_path)
    print(f"Converted: {fname}")

if __name__ == "__main__":
    batch_convert(
        input_path="input/工作規則-easy.docx",
        output_dir="output"
    )
