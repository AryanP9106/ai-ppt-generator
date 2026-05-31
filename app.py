import tkinter as tk
from tkinter import ttk, messagebox
import threading
import sqlite3
import os
import re
import urllib.parse
import requests
import time
import ollama
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor

# ─── CLOUD MODELS ────────────────────────────────────────────────
CLOUD_MODELS = [
    "gemma:2b",       # 1.7 GB  – fast, runs on any machine
    "gemma4:latest",  # 9.6 GB  – best quality, needs ~10 GB RAM free
]

UNSPLASH_KEY = "FBT7VO06efjln1SnF3l_yKjMYP5N7ReS1ti4qzjvig4"

# ─── COLOURS ─────────────────────────────────────────────────────
BG_COLOR     = RGBColor(0x06, 0x5A, 0x82)
ACCENT_COLOR = RGBColor(0x02, 0xC3, 0x9A)
TEXT_LIGHT   = RGBColor(0xFF, 0xFF, 0xFF)
TEXT_BODY    = RGBColor(0xE8, 0xF4, 0xF8)
SLIDE_W      = Inches(13.33)
SLIDE_H      = Inches(7.5)

# ─── PARSER ──────────────────────────────────────────────────────
STOPWORDS = {
    "a","an","the","of","in","on","at","to","for","with","and","or",
    "is","are","was","were","be","been","by","from","this","that","as",
    "its","it","their","show","showing","image","photo","picture",
    "depicting","photorealistic","professional"
}

def parse_slide_response(raw_text):
    title, content_lines, image_prompt, mode = "", [], "", None
    for line in raw_text.strip().splitlines():
        line = line.strip()
        if not line:
            continue
        up = line.upper()
        if up.startswith("TITLE:"):
            title = re.sub(r"[*#]", "", line[6:]).strip()
            mode = "title"
        elif up.startswith("CONTENT:") or up.startswith("BULLETS:"):
            mode = "content"
        elif up.startswith("IMAGE_PROMPT:"):
            image_prompt = line[len("IMAGE_PROMPT:"):].strip()
            mode = "image"
        elif mode == "content":
            clean = re.sub(r"^[-*•]+\s*", "", line)
            clean = re.sub(r"\*\*(.*?)\*\*", r"\1", clean)
            clean = re.sub(r"\*(.*?)\*", r"\1", clean)
            if clean:
                content_lines.append(clean)
        elif mode == "image":
            image_prompt += " " + line

    kw_src = image_prompt.lower() if image_prompt else title.lower()
    words  = re.findall(r"\b[a-z]{3,}\b", kw_src)
    unique = list(dict.fromkeys(w for w in words if w not in STOPWORDS))
    return {
        "title":          title or "Untitled",
        "content":        "\n".join(content_lines),
        "image_prompt":   image_prompt.strip(),
        "image_keywords": " ".join(unique[:5]),
    }

# ─── DB ──────────────────────────────────────────────────────────
def init_db(db_path):
    conn = sqlite3.connect(db_path)
    conn.execute('''CREATE TABLE IF NOT EXISTS slides (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        slide_number INTEGER, title TEXT, content TEXT,
        image_prompt TEXT, image_keywords TEXT, status TEXT DEFAULT "pending"
    )''')
    conn.commit()
    conn.close()

# ─── PPTX BUILDER ────────────────────────────────────────────────
def clean_md(text):
    text = re.sub(r"^#{1,6}\s*", "", text, flags=re.MULTILINE)
    text = re.sub(r"\*\*(.*?)\*\*", r"\1", text)
    text = re.sub(r"\*(.*?)\*",     r"\1", text)
    text = re.sub(r"^[-*•]+\s*",   "",    text, flags=re.MULTILINE)
    return text.strip()

def add_bg(slide):
    bg = slide.shapes.add_shape(1, 0, 0, SLIDE_W, SLIDE_H)
    bg.fill.solid(); bg.fill.fore_color.rgb = BG_COLOR; bg.line.fill.background()
    sp = slide.shapes._spTree
    sp.remove(bg._element); sp.insert(2, bg._element)

def add_title_bar(slide, title):
    bar = slide.shapes.add_shape(1, 0, 0, SLIDE_W, Inches(1.1))
    bar.fill.solid(); bar.fill.fore_color.rgb = ACCENT_COLOR; bar.line.fill.background()
    txb = slide.shapes.add_textbox(Inches(0.4), Inches(0.1), Inches(12.5), Inches(0.9))
    tf  = txb.text_frame; tf.word_wrap = False
    p   = tf.paragraphs[0]; p.text = title.upper()
    p.font.bold = True; p.font.size = Pt(28); p.font.name = "Calibri"
    p.font.color.rgb = TEXT_LIGHT

def add_content(slide, bullets):
    txb = slide.shapes.add_textbox(Inches(0.5), Inches(1.35), Inches(6.2), Inches(5.8))
    tf  = txb.text_frame; tf.word_wrap = True
    first = True
    for bullet in bullets[:6]:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        run = p.add_run(); run.text = "• " + clean_md(bullet)
        p.font.size = Pt(16); p.font.name = "Calibri"
        p.font.color.rgb = TEXT_BODY; p.space_after = Pt(8)

def add_image(slide, img_path):
    try:
        slide.shapes.add_picture(img_path, Inches(7.1), Inches(1.35), Inches(5.8), Inches(5.5))
    except Exception:
        pass

def build_pptx(db_path, output_path):
    conn   = sqlite3.connect(db_path)
    rows   = conn.execute("SELECT id, title, content FROM slides ORDER BY slide_number").fetchall()
    conn.close()
    prs = Presentation()
    prs.slide_width = SLIDE_W; prs.slide_height = SLIDE_H
    blank = prs.slide_layouts[6]
    for db_id, title, content in rows:
        slide = prs.slides.add_slide(blank)
        bullets = [l.strip() for l in (content or "").splitlines() if l.strip()]
        add_bg(slide)
        add_title_bar(slide, title or f"Slide {db_id}")
        add_content(slide, bullets)
        img = os.path.join(os.path.dirname(db_path), "images", f"slide_{db_id}.jpg")
        if os.path.exists(img):
            add_image(slide, img)
    prs.save(output_path)

# ─── IMAGE DOWNLOADER ────────────────────────────────────────────
def download_images(db_path, log):
    img_dir = os.path.join(os.path.dirname(db_path), "images")
    os.makedirs(img_dir, exist_ok=True)
    conn = sqlite3.connect(db_path)
    rows = conn.execute("SELECT id, image_keywords, title FROM slides").fetchall()
    conn.close()
    for db_id, keywords, title in rows:
        path = os.path.join(img_dir, f"slide_{db_id}.jpg")
        if os.path.exists(path) and os.path.getsize(path) > 1000:
            continue
        query = (keywords or title or "technology").strip()
        url   = (f"https://api.unsplash.com/search/photos"
                 f"?query={urllib.parse.quote(query)}&per_page=1"
                 f"&orientation=landscape&client_id={UNSPLASH_KEY}")
        try:
            r    = requests.get(url, timeout=10)
            data = r.json()
            if r.status_code == 200 and data.get("results"):
                img_data = requests.get(data["results"][0]["urls"]["regular"], timeout=15).content
                with open(path, "wb") as f:
                    f.write(img_data)
                log(f"  🖼  Slide {db_id}: image downloaded")
            else:
                log(f"  ⚠  Slide {db_id}: no image found")
            time.sleep(1.5)
        except Exception as e:
            log(f"  ⚠  Slide {db_id} image error: {e}")

# ─── GUI ─────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("AI PPT Generator")
        self.geometry("680x580")
        self.resizable(False, False)
        self.configure(bg="#0f172a")
        self._build_ui()

    def _build_ui(self):
        FONT      = ("Segoe UI", 11)
        FONT_HEAD = ("Segoe UI", 13, "bold")
        FG        = "#e2e8f0"
        ENTRY_BG  = "#1e293b"
        BTN_BG    = "#02c39a"
        BTN_FG    = "#0f172a"

        tk.Label(self, text="🎯  AI PPT Generator", font=("Segoe UI", 17, "bold"),
                 bg="#0f172a", fg=BTN_BG).pack(pady=(22, 4))
        tk.Label(self, text="Powered by Local Ollama Models", font=("Segoe UI", 9),
                 bg="#0f172a", fg="#64748b").pack(pady=(0, 18))

        frm = tk.Frame(self, bg="#0f172a")
        frm.pack(fill="x", padx=40)

        # Topic
        tk.Label(frm, text="Topic", font=FONT_HEAD, bg="#0f172a", fg=FG, anchor="w").grid(
            row=0, column=0, sticky="w", pady=(0, 4))
        self.topic_var = tk.StringVar()
        tk.Entry(frm, textvariable=self.topic_var, font=FONT, bg=ENTRY_BG, fg=FG,
                 insertbackground=FG, relief="flat", bd=6, width=52).grid(
            row=1, column=0, columnspan=2, sticky="ew", pady=(0, 14))

        # Slides + Model row
        tk.Label(frm, text="Number of Slides", font=FONT_HEAD, bg="#0f172a", fg=FG, anchor="w").grid(
            row=2, column=0, sticky="w", pady=(0, 4))
        tk.Label(frm, text="Model", font=FONT_HEAD, bg="#0f172a", fg=FG, anchor="w").grid(
            row=2, column=1, sticky="w", padx=(16, 0), pady=(0, 4))

        self.slides_var = tk.IntVar(value=10)
        slide_spin = tk.Spinbox(frm, from_=3, to=30, textvariable=self.slides_var,
                                font=FONT, bg=ENTRY_BG, fg=FG, relief="flat",
                                buttonbackground=ENTRY_BG, width=6)
        slide_spin.grid(row=3, column=0, sticky="w", pady=(0, 14))

        self.model_var = tk.StringVar(value=CLOUD_MODELS[0])
        model_dd = ttk.Combobox(frm, textvariable=self.model_var, values=CLOUD_MODELS,
                                font=FONT, state="readonly", width=28)
        model_dd.grid(row=3, column=1, sticky="w", padx=(16, 0), pady=(0, 14))

        # Output folder
        tk.Label(frm, text="Output Folder", font=FONT_HEAD, bg="#0f172a", fg=FG, anchor="w").grid(
            row=4, column=0, sticky="w", pady=(0, 4))
        self.out_var = tk.StringVar(value=os.path.expanduser("~/Desktop"))
        tk.Entry(frm, textvariable=self.out_var, font=FONT, bg=ENTRY_BG, fg=FG,
                 insertbackground=FG, relief="flat", bd=6, width=52).grid(
            row=5, column=0, columnspan=2, sticky="ew", pady=(0, 20))

        # Generate button
        self.btn = tk.Button(self, text="✨  Generate Presentation", font=("Segoe UI", 12, "bold"),
                             bg=BTN_BG, fg=BTN_FG, relief="flat", bd=0, padx=20, pady=10,
                             cursor="hand2", command=self._start)
        self.btn.pack(pady=(0, 14))

        # Progress
        self.progress = ttk.Progressbar(self, length=600, mode="determinate")
        self.progress.pack(pady=(0, 8))

        # Log box
        log_frame = tk.Frame(self, bg="#0f172a")
        log_frame.pack(fill="both", expand=True, padx=40, pady=(0, 20))
        self.log_box = tk.Text(log_frame, height=10, font=("Consolas", 9),
                               bg="#1e293b", fg="#94a3b8", relief="flat",
                               state="disabled", wrap="word")
        self.log_box.pack(fill="both", expand=True)

        # Style combobox
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TCombobox", fieldbackground="#1e293b", background="#1e293b",
                        foreground="#e2e8f0", selectbackground="#1e293b")

    def _log(self, msg):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        self.update_idletasks()

    def _start(self):
        topic = self.topic_var.get().strip()
        if not topic:
            messagebox.showwarning("Missing Topic", "Please enter a topic.")
            return
        self.btn.configure(state="disabled", text="Generating…")
        self.log_box.configure(state="normal"); self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")
        self.progress["value"] = 0
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self):
        topic      = self.topic_var.get().strip()
        num_slides = self.slides_var.get()
        model      = self.model_var.get()
        out_folder = self.out_var.get().strip()
        os.makedirs(out_folder, exist_ok=True)

        safe_topic = re.sub(r"[^\w\s-]", "", topic).strip().replace(" ", "_")[:40]
        db_path    = os.path.join(out_folder, f"{safe_topic}.db")
        pptx_path  = os.path.join(out_folder, f"{safe_topic}.pptx")

        try:
            # ── Step 1: Init DB ──────────────────────────────────
            init_db(db_path)
            conn = sqlite3.connect(db_path)
            conn.execute("DELETE FROM slides"); conn.commit(); conn.close()

            self._log(f"📋 Topic   : {topic}")
            self._log(f"🤖 Model   : {model}")
            self._log(f"📊 Slides  : {num_slides}\n")

            client = ollama.Client(host="http://localhost:11434")
            total_steps = num_slides + num_slides + 1   # generate + images + build

            # ── Step 2: Generate slides ──────────────────────────
            for i in range(1, num_slides + 1):
                self._log(f"✍  Generating slide {i}/{num_slides}…")
                prompt = (
                    f"You are creating slide {i} of {num_slides} for a presentation titled: \"{topic}\".\n"
                    "Reply in EXACTLY this format — no extra text, no markdown outside the format:\n\n"
                    "TITLE: <concise slide title>\n"
                    "CONTENT:\n"
                    "- <bullet point 1>\n"
                    "- <bullet point 2>\n"
                    "- <bullet point 3>\n"
                    "- <bullet point 4>\n"
                    "IMAGE_PROMPT: <vivid 8-word description of a real photograph that suits this slide>\n"
                )
                resp   = client.generate(model=model, prompt=prompt)
                raw    = getattr(resp, "response", str(resp)).strip()
                parsed = parse_slide_response(raw)

                conn = sqlite3.connect(db_path)
                conn.execute(
                    "INSERT INTO slides (slide_number,title,content,image_prompt,image_keywords,status) "
                    "VALUES (?,?,?,?,?,?)",
                    (i, parsed["title"], parsed["content"],
                     parsed["image_prompt"], parsed["image_keywords"], "pending")
                )
                conn.commit(); conn.close()
                self._log(f"   ✅ {parsed['title']}")
                self.progress["value"] = (i / total_steps) * 100

            # ── Step 3: Download images ──────────────────────────
            self._log("\n🖼  Downloading images from Unsplash…")
            download_images(db_path, self._log)
            self.progress["value"] = ((num_slides + num_slides) / total_steps) * 100

            # ── Step 4: Build PPTX ───────────────────────────────
            self._log("\n🔨 Building presentation…")
            build_pptx(db_path, pptx_path)
            self.progress["value"] = 100

            self._log(f"\n✅ Done! Saved to:\n   {pptx_path}")
            messagebox.showinfo("Done!", f"Presentation saved to:\n{pptx_path}")
            os.startfile(pptx_path)   # auto-open on Windows

        except Exception as e:
            self._log(f"\n❌ Error: {e}")
            messagebox.showerror("Error", str(e))
        finally:
            self.btn.configure(state="normal", text="✨  Generate Presentation")


if __name__ == "__main__":
    App().mainloop()
