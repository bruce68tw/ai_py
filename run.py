from pdf_article import pdf_article
from extract_text import extract_lines
from structure_parser import parse_structure
from chunk_builder import build_chunks
from llm_extract import extract_candidates
from validator import validate

##pip install pdfplumber
#pip install pdfplumber camelot-py pandas
#pip install Pillow
#pip install tqdm   #看進度
#pip install python-dotenv  #串接 LLM API

PDF_PATH = "input/工作規則.pdf"

def main():
    lines = pdf_article(PDF_PATH)
    #lines = extract_lines(PDF_PATH)
    #articles = parse_structure(lines)
    #chunks = build_chunks(articles)
    #candidates = extract_candidates(chunks)
    #validate(candidates)

if __name__ == "__main__":
    main()
