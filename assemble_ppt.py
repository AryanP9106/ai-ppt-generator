import sqlite3
import os
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor

def build_presentation():
    db_path = 'presentation.db'
    if not os.path.exists(db_path):
        print("❌ Error: presentation.db not found!")
        return

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # Check if the table and columns exist
    try:
        cursor.execute("SELECT id, title, content FROM slides")
        rows = cursor.fetchall()
    except sqlite3.OperationalError as e:
        print(f"❌ Database Error: {e}")
        return

    if not rows:
        print("⚠️ Database is empty. Re-run your populate script.")
        return

    prs = Presentation()
    # Set to 16:9 Widescreen
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    print(f"Starting assembly of {len(rows)} slides...")

    for row in rows:
        db_id, title, content = row
        print(f"-> Processing Slide {db_id}")

        # Add a blank slide
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # 1. THE TITLE
        left, top, width, height = Inches(0.5), Inches(0.3), Inches(12), Inches(1)
        txt_title = slide.shapes.add_textbox(left, top, width, height)
        tf_t = txt_title.text_frame
        p_t = tf_t.paragraphs[0]
        p_t.text = str(title).replace("TITLE:", "").strip().upper()
        p_t.font.bold = True
        p_t.font.size = Pt(36)
        p_t.font.name = 'Arial'

        # 2. THE IMAGE (Right Side)
        img_path = f"images/slide_{db_id}.jpg"
        if os.path.exists(img_path):
            try:
                # Inches(7) from left makes it sit on the right half
                slide.shapes.add_picture(img_path, Inches(7), Inches(1.5), width=Inches(5.8))
            except Exception as e:
                print(f"   Skipping image {db_id} due to file error.")

        # 3. THE CONTENT (Left Side)
        left_c, top_c, width_c, height_c = Inches(0.5), Inches(1.5), Inches(6), Inches(5)
        txt_content = slide.shapes.add_textbox(left_c, top_c, width_c, height_c)
        tf_c = txt_content.text_frame
        tf_c.word_wrap = True

        # Split AI response into clean lines
        lines = [line.strip() for line in str(content).split('\n') if line.strip()]
        for line in lines:
            if "CONTENT:" in line.upper(): continue
            p = tf_c.add_paragraph()
            p.text = line.lstrip('-*• ').strip()
            p.font.size = Pt(20)
            p.space_after = Pt(10)

    output = "Final_Clean_Presentation.pptx"
    try:
        prs.save(output)
        print(f"\n✅ SUCCESS: {output} created!")
    except PermissionError:
        print(f"\n❌ ERROR: Close {output} before running the script!")

    conn.close()

if __name__ == "__main__":
    build_presentation()