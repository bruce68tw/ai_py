import json
import os

def write_debug(step: str, data, debug_dir="debug"):
    os.makedirs(debug_dir, exist_ok=True)
    path = os.path.join(debug_dir, f"{step}.json")

    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"[DEBUG] write {path}")
