# Qwen3 sometimes thinks out loud and write some explanation, we need to extract only JSON from his answer
def clear_qwen_text(raw_text):
    # Remove <think>...</think> block first
    if "<think>" in raw_text and "</think>" in raw_text:
        raw_text = raw_text.split("</think>")[-1]

    # Clean markdown
    clean_text = raw_text.strip()
    if clean_text.startswith("```") and clean_text.endswith("```"):
        clean_text = clean_text.split("```")[1]
        if clean_text.startswith("json"):
            clean_text = clean_text[4:]
        clean_text = clean_text.rsplit("```", 1)[0]

    return clean_text