import sqlite3
import requests
import os
import time

# REPLACE WITH YOUR UNSPLASH ACCESS KEY
ACCESS_KEY = "FBT7VO06efjln1SnF3l_yKjMYP5N7ReS1ti4qzjvig4"

def download_unsplash_images():
    if not os.path.exists('images'):
        os.makedirs('images')

    conn = sqlite3.connect('presentation.db')
    cursor = conn.cursor()
    cursor.execute("SELECT id, title FROM slides")
    rows = cursor.fetchall()

    print(f"Fetching {len(rows)} professional photos from Unsplash...")

    for row in rows:
        db_id, title = row
        image_path = f"images/slide_{db_id}.jpg"

        if os.path.exists(image_path) and os.path.getsize(image_path) > 1000:
            continue

        # Search for a professional photo matching the slide title
        search_url = f"https://api.unsplash.com/search/photos?query={title}&per_page=1&client_id={ACCESS_KEY}"
        
        try:
            print(f"Searching for {title}...", end=" ")
            response = requests.get(search_url, timeout=10)
            data = response.json()

            if response.status_code == 200 and data['results']:
                img_url = data['results'][0]['urls']['regular']
                img_data = requests.get(img_url).content
                with open(image_path, 'wb') as f:
                    f.write(img_data)
                print("✅ Done")
            else:
                print("❌ No photo found.")
            
            # Unsplash free tier allows 50 requests per hour. 
            # We wait 2 seconds to be safe.
            time.sleep(2)

        except Exception as e:
            print(f"❌ Error: {e}")
            time.sleep(5)

    conn.close()

if __name__ == "__main__":
    download_unsplash_images()