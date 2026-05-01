import sqlite3
import ollama

def generate_and_save_slides(topic, num_slides):
    # Connect to DB
    conn = sqlite3.connect('presentation.db')
    cursor = conn.cursor()

    # Create a client with a longer timeout (in case the model is slow to load)
    client = ollama.Client(host='http://localhost:11434')

    for i in range(1, num_slides + 1):
        print(f"--- Requesting Slide {i} from Gemma 4... Please wait ---")
        
        prompt = (
            f"Topic: {topic}. Generate content for slide {i}. "
            "Format your response exactly like this:\n"
            "TITLE: Your Slide Title\n"
            "CONTENT:\n"
            "- First bullet point\n"
            "- Second bullet point\n"
            "IMAGE_PROMPT: A photorealistic description for an image."
        )

        try:
            response = client.generate(model="gemma4", prompt=prompt)
            content = str(response).strip()

            cursor.execute('''
                INSERT INTO slides (slide_number, title, content, status) 
                VALUES (?, ?, ?, ?)
            ''', (i, f"Slide {i}", content, 'pending'))

            conn.commit() # Save after every slide
            print(f"DONE: Slide {i} saved to database.")
        except Exception as e:
            print(f"ERROR on Slide {i}: {e}")
            break

    conn.close()
    print("\nProcessing complete.")

# Let's test with just 2 slides first to ensure the connection is solid
generate_and_save_slides("Future of AI in Software Engineering", 60)