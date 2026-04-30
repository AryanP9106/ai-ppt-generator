import sqlite3
import os
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.text.text import TextFrame
from typing import cast

def build_from_db():
    conn = sqlite3.connect('presentation.db')
    cursor = conn.cursor()
    cursor.execute("SELECT title, content, image_prompt FROM slides")
    rows = cursor.fetchall()

    if not rows:
        print("No data found.")
        return

    prs = Presentation()

    for row in rows:
        db_title, db_content, db_img_prompt = row
        
        # Use 'Title and Content' layout
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)

        # 1. SET TITLE
        try:
            if slide.shapes.title:
                slide.shapes.title.text = str(db_title).replace("TITLE:", "").strip()
        except:
            pass # Skip if title shape is weird

        # 2. SET CONTENT
        for shape in slide.placeholders:
            try:
                if shape.placeholder_format.type == PP_PLACEHOLDER.BODY:
                    tf = getattr(shape, "text_frame", None)
                    if tf is None:
                        continue

                    # or: tf = cast(TextFrame, shape.text_frame)

                    tf.clear()
                    tf.word_wrap = True

                    lines = str(db_content).split('\n')
                    for line in lines:
                        clean = line.strip()
                        if not clean or "CONTENT:" in clean.upper():
                            continue

                        p = tf.add_paragraph()
                        p.text = clean.lstrip('-*• ').strip()
                        p.font.size = Pt(20)
                        p.level = 0
            except:
                continue # Skip shapes that don't support text_frame

        # 3. PLACEHOLDER FOR IMAGES (The 'Superb' part)
        # For now, we will just print where the image would go
        # Later, we will download them using a Free API
        if db_img_prompt:
            print(f"Slide prepared for image: {db_img_prompt[:30]}...")

    output_name = "Final_Presentation_Robust.pptx"
    try:
        prs.save(output_name)
        print(f"\nSUCCESS: Generated {output_name} with {len(rows)} slides.")
    except PermissionError:
        print("\nERROR: Close the PPT file before running!")

    conn.close()

if __name__ == "__main__":
    build_from_db()