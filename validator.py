from utils.debug_writer import write_debug

def validate(candidates):
    validated = []

    for item in candidates:
        validated.append({
            "chunk_id": item["chunk_id"],
            "status": "pending",
            "reason": "not validated yet",
            "data": item["llm_output"]
        })

    write_debug("04_validated", validated)
    return validated
