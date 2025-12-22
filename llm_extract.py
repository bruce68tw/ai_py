from utils.debug_writer import write_debug

def extract_candidates(chunks):
    results = []

    for c in chunks:
        # 先假設你之後會接真正 LLM
        llm_output = {
            "entities": [],
            "relations": [],
            "constraints": [],
            "exceptions": []
        }

        results.append({
            "chunk_id": c["chunk_id"],
            "input": c["text"],
            "llm_output": llm_output
        })

    write_debug("03_llm_candidates", results)
    return results
