import os
import sys
import shutil
import sqlite3
import threading
import time
import zipfile
from concurrent.futures import ThreadPoolExecutor
from multiprocessing import Process

import pytesseract
from PIL import Image, ImageGrab, ImageTk  # Utilizziamo Image.Resampling.LANCZOS
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import tkinter as tk
from pynput import keyboard

# --- Configurazione --- #
ASSETS_DIR = "assets"
IMAGES_DIR = "images"
DB_FILE = "images_text.db"

ALLOWED_IMAGE_EXT = {".png", ".jpg", ".jpeg", ".bmp", ".tiff"}
ALLOWED_DOC_EXT = {".docx"}
ALLOWED_PPT_EXT = {".pptx"}

MAX_THREADS = 4

# --- Funzioni di utilità --- #
def ensure_directories():
    if not os.path.exists(ASSETS_DIR):
        os.makedirs(ASSETS_DIR)
    if not os.path.exists(IMAGES_DIR):
        os.makedirs(IMAGES_DIR)

def get_next_image_id():
    max_id = 0
    for fname in os.listdir(IMAGES_DIR):
        if fname.startswith("img_"):
            try:
                num = int(fname.split("_")[1].split(".")[0])
                if num > max_id:
                    max_id = num
            except Exception:
                continue
    return max_id + 1

def save_image(source_path, image_bytes=None, ext=None):
    global_lock = threading.Lock()
    with global_lock:
        new_id = get_next_image_id()
    if image_bytes is None:
        ext = os.path.splitext(source_path)[1].lower()
    else:
        if not ext:
            ext = ".png"
    new_filename = f"img_{new_id}{ext}"
    dest_path = os.path.join(IMAGES_DIR, new_filename)
    if image_bytes is None:
        shutil.move(source_path, dest_path)
    else:
        with open(dest_path, "wb") as f:
            f.write(image_bytes)
    print(f"[INFO] Immagine salvata: {dest_path}")
    return dest_path

# --- Estrazione immagini da documenti --- #
def extract_images_from_docx(docx_path):
    try:
        with zipfile.ZipFile(docx_path, "r") as z:
            for file_info in z.infolist():
                if file_info.filename.startswith("word/media/"):
                    fname = file_info.filename.split("/")[-1]
                    ext = os.path.splitext(fname)[1].lower()
                    if ext in ALLOWED_IMAGE_EXT:
                        image_bytes = z.read(file_info.filename)
                        save_image(docx_path, image_bytes=image_bytes, ext=ext)
        print(f"[INFO] Estrazione immagini da {docx_path} completata.")
    except Exception as e:
        print(f"[ERRORE] Errore nell'estrazione da DOCX {docx_path}: {e}")

def extract_images_from_pptx(pptx_path):
    try:
        with zipfile.ZipFile(pptx_path, "r") as z:
            for file_info in z.infolist():
                if file_info.filename.startswith("ppt/media/"):
                    fname = file_info.filename.split("/")[-1]
                    ext = os.path.splitext(fname)[1].lower()
                    if ext in ALLOWED_IMAGE_EXT:
                        image_bytes = z.read(file_info.filename)
                        save_image(pptx_path, image_bytes=image_bytes, ext=ext)
        print(f"[INFO] Estrazione immagini da {pptx_path} completata.")
    except Exception as e:
        print(f"[ERRORE] Errore nell'estrazione da PPTX {pptx_path}: {e}")

def process_file(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    if ext in ALLOWED_IMAGE_EXT:
        print(f"[INFO] Elaborazione file immagine: {file_path}")
        save_image(file_path)
    elif ext in ALLOWED_DOC_EXT:
        print(f"[INFO] Elaborazione file DOCX: {file_path}")
        extract_images_from_docx(file_path)
        processed_dir = os.path.join(ASSETS_DIR, "processed")
        if not os.path.exists(processed_dir):
            os.makedirs(processed_dir)
        shutil.move(file_path, os.path.join(processed_dir, os.path.basename(file_path)))
    elif ext in ALLOWED_PPT_EXT:
        print(f"[INFO] Elaborazione file PPTX: {file_path}")
        extract_images_from_pptx(file_path)
        processed_dir = os.path.join(ASSETS_DIR, "processed")
        if not os.path.exists(processed_dir):
            os.makedirs(processed_dir)
        shutil.move(file_path, os.path.join(processed_dir, os.path.basename(file_path)))
    else:
        print(f"[WARNING] Formato non supportato per il file: {file_path}")

def process_existing_assets():
    for entry in os.listdir(ASSETS_DIR):
        file_path = os.path.join(ASSETS_DIR, entry)
        if os.path.isfile(file_path):
            process_file(file_path)

# --- Gestione del database --- #
def setup_db():
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS images_text (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT UNIQUE,
            text TEXT
        )
    """)
    conn.commit()
    conn.close()

def save_text_in_db(filename, text):
    try:
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute("SELECT 1 FROM images_text WHERE text = ?", (text,))
        if cur.fetchone():
            print(f"[INFO] Testo già presente nel DB, non viene inserito per {filename}.")
            conn.close()
            return
        cur.execute("INSERT INTO images_text (filename, text) VALUES (?, ?)", (filename, text))
        conn.commit()
        conn.close()
        print(f"[INFO] Testo salvato per {filename}")
    except Exception as e:
        print(f"[ERRORE] Errore nel salvataggio su DB: {e}")

def search_in_db(query, interactive=True):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("SELECT filename, text FROM images_text WHERE text LIKE ?", (f"%{query}%",))
    results = cur.fetchall()
    conn.close()
    if results:
        print("=== Risultati ricerca ===")
        for i, (filename, text) in enumerate(results, start=1):
            print(f"{i}. File: {filename}\n   Testo estratto (inizio): {text[:100]}...\n")
        if interactive and sys.stdin.isatty():
            selection = input("Inserisci il numero del risultato per aprire l'immagine (oppure premi invio per saltare): ")
            if selection.strip():
                try:
                    idx = int(selection.strip()) - 1
                    if 0 <= idx < len(results):
                        filename, _ = results[idx]
                        image_path = os.path.join(IMAGES_DIR, filename)
                        print(f"[INFO] Apertura immagine: {image_path}")
                        img = Image.open(image_path)
                        img.show()
                    else:
                        print("[WARNING] Numero non valido.")
                except ValueError:
                    print("[WARNING] Input non valido.")
    else:
        print("Nessun risultato trovato.")

# --- Elaborazione OCR --- #
def process_image(image_path):
    try:
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute("SELECT 1 FROM images_text WHERE filename = ?", (os.path.basename(image_path),))
        if cur.fetchone():
            print(f"[INFO] {image_path} già processato.")
            conn.close()
            return
        conn.close()
        img = Image.open(image_path)
        text = pytesseract.image_to_string(img, lang="ita")
        save_text_in_db(os.path.basename(image_path), text)
    except Exception as e:
        print(f"[ERRORE] Errore nell'elaborazione OCR di {image_path}: {e}")

def process_all_images():
    image_files = [os.path.join(IMAGES_DIR, f) for f in os.listdir(IMAGES_DIR)
                   if os.path.splitext(f)[1].lower() in ALLOWED_IMAGE_EXT]
    if not image_files:
        print("[INFO] Nessuna immagine da elaborare nella cartella images.")
        return
    with ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
        executor.map(process_image, image_files)

# --- Watchdog per la cartella assets --- #
class AssetsHandler(FileSystemEventHandler):
    def on_created(self, event):
        if not event.is_directory:
            time.sleep(1)
            print(f"[INFO] Nuovo file rilevato: {event.src_path}")
            process_file(event.src_path)
            process_all_images()

# --- Modalità Screen Grabber (selezione libera) --- #
class ScreenGrabber(tk.Tk):
    def __init__(self):
        super().__init__()
        self.tk.call('tk', 'scaling', 1.0)
        self.overrideredirect(True)
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        self.geometry(f"{screen_width}x{screen_height}+0+0")
        screenshot = ImageGrab.grab()
        screenshot = screenshot.resize((screen_width, screen_height), resample=Image.Resampling.LANCZOS)
        self.screenshot = screenshot
        self.bg_image = ImageTk.PhotoImage(screenshot)
        self.canvas = tk.Canvas(self, width=screen_width, height=screen_height)
        self.canvas.pack(fill="both", expand=True)
        self.canvas.create_image(0, 0, image=self.bg_image, anchor="nw")
        self.start_x = self.start_y = None
        self.rect = None
        self.selected_area = None
        self.bind("<ButtonPress-1>", self.on_button_press)
        self.bind("<B1-Motion>", self.on_move_press)
        self.bind("<ButtonRelease-1>", self.on_button_release)
        self.bind("<Escape>", lambda event: self.destroy())
        self.focus_force()

    def on_button_press(self, event):
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
        self.rect = self.canvas.create_rectangle(
            self.start_x, self.start_y, self.start_x, self.start_y,
            outline='red', width=3, fill=''
        )

    def on_move_press(self, event):
        curX, curY = (self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))
        self.canvas.coords(self.rect, self.start_x, self.start_y, curX, curY)

    def on_button_release(self, event):
        end_x = self.canvas.canvasx(event.x)
        end_y = self.canvas.canvasy(event.y)
        self.selected_area = (int(min(self.start_x, end_x)),
                              int(min(self.start_y, end_y)),
                              int(max(self.start_x, end_x)),
                              int(max(self.start_y, end_y)))
        self.destroy()

def image_grabber():
    grabber = ScreenGrabber()
    grabber.mainloop()
    if grabber.selected_area:
        print(f"[INFO] Area selezionata: {grabber.selected_area}")
        img = ImageGrab.grab(bbox=grabber.selected_area)
        text = pytesseract.image_to_string(img, lang="ita")
        print("=== Testo estratto dall'area ===")
        print(text)
        print("=== Risultati ricerca nel DB ===")
        search_in_db(text, interactive=False)
    else:
        print("[INFO] Nessuna area selezionata.")

# --- Modalità Text Box Selector --- #
def text_box_selector():
    root = tk.Tk()
    root.overrideredirect(True)
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    root.geometry(f"{screen_width}x{screen_height}+0+0")
    # Se lo schermo è molto grande, usa una scala moderata; altrimenti 1.0
    scale = 1.0 if screen_width <= 1600 else 0.75
    new_width = int(screen_width * scale)
    new_height = int(screen_height * scale)
    screenshot = ImageGrab.grab()
    screenshot_scaled = screenshot.resize((new_width, new_height), resample=Image.Resampling.LANCZOS)
    bg_image = ImageTk.PhotoImage(screenshot_scaled)
    canvas = tk.Canvas(root, width=new_width, height=new_height)
    canvas.pack(fill="both", expand=True)
    canvas.create_image(0, 0, image=bg_image, anchor="nw")
    # Usa pytesseract per ottenere i dati a livello 2 (block) per box più ampi
    data = pytesseract.image_to_data(screenshot_scaled, lang="ita", output_type=pytesseract.Output.DICT)
    boxes = []
    for i in range(len(data['level'])):
        if data['level'][i] == 2:  # livello block
            text_val = data['text'][i].strip()
            try:
                conf = float(data['conf'][i])
            except:
                conf = 0
            if text_val and conf > 0:
                x = data['left'][i]
                y = data['top'][i]
                w = data['width'][i]
                h = data['height'][i]
                boxes.append((x, y, x+w, y+h))
    box_items = []
    for box in boxes:
        x1, y1, x2, y2 = box
        item = canvas.create_rectangle(x1, y1, x2, y2, outline='blue', width=2)
        box_items.append((item, box))
    def on_click(event):
        for item, box in box_items:
            x1, y1, x2, y2 = box
            if x1 <= event.x <= x2 and y1 <= event.y <= y2:
                cropped = screenshot_scaled.crop(box)
                ocr_text = pytesseract.image_to_string(cropped, lang="ita")
                print("=== Testo estratto dal blocco selezionato ===")
                print(ocr_text)
                cropped.show()
                root.destroy()
                return
    canvas.bind("<Button-1>", on_click)
    root.bind("<Escape>", lambda e: root.destroy())
    root.focus_force()
    root.mainloop()

def on_activate_text_box_selector():
    print("[INFO] Attivazione Text Box Selector tramite hotkey")
    p = Process(target=text_box_selector)
    p.start()

# --- Registrazione hotkey tramite pynput --- #
def on_activate_image_grabber():
    print("[INFO] Attivazione Image Grabber tramite hotkey")
    p = Process(target=image_grabber)
    p.start()

hotkey_listener = keyboard.GlobalHotKeys({
    '<ctrl>+<shift>+g': on_activate_image_grabber,
    '<ctrl>+<shift>+s': on_activate_text_box_selector
})
hotkey_listener.start()

# --- Funzione principale --- #
def main():
    ensure_directories()
    setup_db()
    process_existing_assets()
    event_handler = AssetsHandler()
    observer = Observer()
    observer.schedule(event_handler, ASSETS_DIR, recursive=False)
    observer.start()
    print(f"[INFO] Inizio monitoraggio della cartella '{ASSETS_DIR}'...")
    process_all_images()
    print("[INFO] Premi Ctrl+Shift+G per attivare l'Image Grabber.")
    print("[INFO] Premi Ctrl+Shift+S per attivare il Text Box Selector.")
    try:
        while True:
            query = input("Inserisci una frase per cercare nel DB (oppure 'exit' per uscire): ")
            if query.lower() == "exit":
                break
            search_in_db(query)
    except KeyboardInterrupt:
        pass
    finally:
        observer.stop()
        observer.join()
        print("[INFO] Programma terminato.")

if __name__ == "__main__":
    main()
