from utils.debug_writer import write_debug

def build_chunks(articles):
    chunks = []

    for art in articles:
        for idx, para in enumerate(art["paragraphs"], start=1):
            chunks.append({
                "chunk_id": f"{art['article']}_p{idx}",
                "article": art["article"],
                "text": para,
                "page": art["page"]
            })

    write_debug("02_chunks", chunks)
    return chunks
