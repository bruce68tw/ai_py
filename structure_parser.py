import re
from utils.debug_writer import write_debug

ARTICLE_PATTERN = re.compile(r"^第\s*[一二三四五六七八九十百千0-9]+\s*條")

def parse_structure(lines):
    articles = []
    current = None

    for item in lines:
        text = item["text"]

        if ARTICLE_PATTERN.match(text):
            current = {
                "article": text,
                "page": item["page"],
                "paragraphs": []
            }
            articles.append(current)
        else:
            if current:
                current["paragraphs"].append(text)

    write_debug("01_articles", articles)
    return articles
