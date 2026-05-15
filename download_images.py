import sqlite3
import requests
import os
import time
import urllib.parse

ACCESS_KEY = "FBT7VO06efjln1SnF3l_yKjMYP5N7ReS1ti4qzjvig4"

def download_unsplash_images(db_path="presentation.db"):
    os.makedirs("images", exist_ok=True)

    conn   = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # Use image_keywords (derived from IMAGE_PROMPT) if available,
    # fall back to title if the column doesn't exist yet
    try:
        cursor.execute("SELECT id, image_keywords, title FROM slides")
    except sqlite3.OperationalError:
        cursor.execute("SELECT id, title, title FROM slides")

    rows = cursor.fetchall()
    print(f"Fetching images for {len(rows)} slides...\n")

    for db_id, keywords, title in rows:
        image_path = f"images/slide_{db_id}.jpg"

        if os.path.exists(image_path) and os.path.getsize(image_path) > 1000:
            print(f"  Slide {db_id}: already downloaded, skipping.")
            continue

        # Prefer keywords from IMAGE_PROMPT; fall back to title
        query = (keywords or title or "technology").strip()
        query_enc = urllib.parse.quote(query)

        url = (
            f"https://api.unsplash.com/search/photos"
            f"?query={query_enc}&per_page=1&orientation=landscape"
            f"&client_id={ACCESS_KEY}"
        )

        try:
            print(f"  Slide {db_id}: searching \"{query}\"...", end=" ", flush=True)
            r    = requests.get(url, timeout=10)
            data = r.json()

            if r.status_code == 200 and data.get("results"):
                img_url  = data["results"][0]["urls"]["regular"]
                img_data = requests.get(img_url, timeout=15).content
                with open(image_path, "wb") as f:
                    f.write(img_data)
                print("✅")
            else:
                print(f"❌ ({r.status_code} – {data.get('errors', 'no results')})")

            time.sleep(2)   # respect Unsplash free-tier rate limit

        except Exception as e:
            print(f"❌ Error: {e}")
            time.sleep(5)

    conn.close()
    print("\nImage download complete.")


if __name__ == "__main__":
    download_unsplash_images()
