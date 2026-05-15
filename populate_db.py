import sqlite3
import re
import ollama
from database import init_db

# ──────────────────────────────────────────────
# Parse the structured AI response into fields
# ──────────────────────────────────────────────
def parse_slide_response(raw_text):
    title          = ""
    content_lines  = []
    image_prompt   = ""
    mode           = None

    for line in raw_text.strip().splitlines():
        line = line.strip()
        if not line:
            continue

        up = line.upper()

        if up.startswith("TITLE:"):
            title = re.sub(r"[*#]", "", line[6:]).strip()
            mode  = "title"

        elif up.startswith("CONTENT:") or up.startswith("BULLETS:"):
            mode = "content"

        elif up.startswith("IMAGE_PROMPT:"):
            image_prompt = line[len("IMAGE_PROMPT:"):].strip()
            mode = "image"

        elif mode == "content":
            clean = re.sub(r"^[-*•]+\s*", "", line)
            clean = re.sub(r"\*\*(.*?)\*\*", r"\1", clean)
            clean = re.sub(r"\*(.*?)\*",   r"\1", clean)
            if clean:
                content_lines.append(clean)

        elif mode == "image":
            image_prompt += " " + line

    # Build focused keywords for Unsplash from the IMAGE_PROMPT, not the title
    stopwords = {
        "a","an","the","of","in","on","at","to","for","with","and",
        "or","is","are","was","were","be","been","by","from","this",
        "that","as","its","it","their","show","showing","image","photo",
        "picture","depicting","photorealistic","professional"
    }
    kw_src  = image_prompt.lower() if image_prompt else title.lower()
    words   = re.findall(r"\b[a-z]{3,}\b", kw_src)
    unique  = list(dict.fromkeys(w for w in words if w not in stopwords))
    keywords = " ".join(unique[:5])   # top 5 distinct words

    return {
        "title":          title or "Untitled",
        "content":        "\n".join(content_lines),
        "image_prompt":   image_prompt.strip(),
        "image_keywords": keywords,
    }


# ──────────────────────────────────────────────
# Main generator
# ──────────────────────────────────────────────
def generate_and_save_slides(topic, num_slides, model="gemma4", db_path="presentation.db"):
    init_db(db_path)

    conn   = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM slides")   # fresh run
    conn.commit()
    print(f"Generating {num_slides} slides on: {topic}\n")

    client = ollama.Client(host="http://localhost:11434")

    for i in range(1, num_slides + 1):
        print(f"--- Slide {i}/{num_slides} ---")

        prompt = (
            f"You are creating slide {i} of {num_slides} for a presentation on: {topic}.\n"
            "Reply in EXACTLY this format and nothing else:\n\n"
            "TITLE: <concise slide title>\n"
            "CONTENT:\n"
            "- <bullet 1>\n"
            "- <bullet 2>\n"
            "- <bullet 3>\n"
            "IMAGE_PROMPT: <vivid 10-word description of a real-world photo that fits this slide>\n"
        )

        try:
            resp    = client.generate(model=model, prompt=prompt)
            # .response is the plain text string
            raw     = getattr(resp, "response", str(resp)).strip()
            parsed  = parse_slide_response(raw)

            cursor.execute(
                "INSERT INTO slides (slide_number, title, content, image_prompt, image_keywords, status) "
                "VALUES (?, ?, ?, ?, ?, ?)",
                (i, parsed["title"], parsed["content"],
                 parsed["image_prompt"], parsed["image_keywords"], "pending")
            )
            conn.commit()
            print(f"  Title   : {parsed['title']}")
            print(f"  Keywords: {parsed['image_keywords']}\n")

        except Exception as e:
            print(f"  ERROR: {e}\n")
            break

    conn.close()
    print("Done.")


if __name__ == "__main__":
    generate_and_save_slides("Future of AI in Software Engineering", 10)
