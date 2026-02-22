"""
NWG-Bericht Converter - Clean Version
=====================================

Dieses Tool konvertiert Excel-Daten in Word-Berichte durch Ersetzen von Content Controls.

Hauptfunktionen:
- Excel-Import mit Drag & Drop
- Berater-Auswahl aus Datenbank
- Word Content Control Ersetzung
- Tag-Validation (zeigt fehlende Werte an)
- Moderne GUI mit leicht abgerundeten Buttons
- Easter Egg (Doppelklick auf Logo)

Entwickler: Elia Salemi
Version: Clean (70% weniger Code)
"""

import os
import math
import logging
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from docx import Document
from docx.oxml.ns import qn
from PIL import Image, ImageTk
from tkinterdnd2 import TkinterDnD, DND_FILES
import sys
import getpass
from pathlib import Path
import logging.handlers

# ========== Pfad zum externen Vorlagen-Ordner ==========
def get_vorlagen_path():
    """Gibt den Pfad zum externen Vorlagen-Ordner zurück"""
    if getattr(sys, 'frozen', False):
        # Wenn als .exe gestartet → Ordner neben der .exe
        base_path = Path(sys.executable).parent
    else:
        # Wenn als Skript ausgeführt → Projektordner
        base_path = Path(__file__).parent

    vorlagen_dir = base_path / "Vorlagen"
    if not vorlagen_dir.exists():
        print(f"⚠️  Vorlagen-Ordner nicht gefunden: {vorlagen_dir}")
    return vorlagen_dir

VORLAGEN_PATH = get_vorlagen_path()

# ========== Pfade & Konfiguration ==========
def get_resource_path(relative_path):
    """Ressourcen-Pfad für .exe und Entwicklung"""
    try:
        return os.path.join(sys._MEIPASS, relative_path)
    except AttributeError:
        return os.path.join(os.path.dirname(__file__), "Vorlagen", relative_path)

# Pfade automatisch bestimmen
BASE_DIR = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else str(Path(__file__).parent.parent)
BERATER_LISTE = str(VORLAGEN_PATH / "Energieberaterliste_T2.xlsx")
LOGO_PATH = get_resource_path("logo.jpg")
ICON_PATH = get_resource_path("Converter_logo.ico")

# Logging-Setup
logs_dir = os.path.join(BASE_DIR, "Logs")
os.makedirs(logs_dir, exist_ok=True)
log_file = os.path.join(logs_dir, f"converter_{getpass.getuser()}.log")

# Beim Start: bestehende Datei → .log.1 (jede Session bekommt eigene Datei)
_log_handler = logging.handlers.RotatingFileHandler(log_file, maxBytes=20*1024, backupCount=3)
if os.path.exists(log_file):
    _log_handler.doRollover()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s: %(message)s",
    handlers=[_log_handler]
)

# ========== GUI Konstanten ==========
COLORS = {
    'primary': "#20A065",
    'secondary': "#70AD47", 
    'accent': "#B0CD5C",
    'background': "#ffffff",
    'text': "#403f45",
    'drop_bg': "#eafaf1"
}

FONTS = {
    'label': ("Arial", 10),
    'button': ("Arial", 10, "bold"),
    'header': ('Helvetica', 18, 'bold')
}

# ========== Globale Variablen ==========
excel_datei = None
berater_df = pd.DataFrame()
werte_dict = {}
WORD_NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

# ========== Moderne Buttons ==========
class ModernButton(tk.Canvas):
    """Moderne Buttons mit leichter Rundung und Hover-Effekt"""
    
    def __init__(self, parent, text, command, width=200, height=44, **kwargs):
        bg_color = kwargs.get('bg', COLORS['primary'])
        super().__init__(parent, width=width, height=height, bg=parent['bg'], highlightthickness=0)
        
        self.command = command
        self.bg_color = bg_color
        self.hover_color = kwargs.get('hover_bg', COLORS['secondary'])
        
        # Leicht abgerundetes Rechteck (12px Radius)
        self.rect = self.create_rounded_rect(2, 2, width-2, height-2, radius=12, 
                                           fill=bg_color, outline=bg_color)
        self.text_id = self.create_text(width//2, height//2, text=text, 
                                      fill="white", font=FONTS['button'])
        
        # Events
        self.bind("<Button-1>", lambda e: command())
        self.bind("<Enter>", lambda e: self.itemconfig(self.rect, fill=self.hover_color))
        self.bind("<Leave>", lambda e: self.itemconfig(self.rect, fill=self.bg_color))
    
    def create_rounded_rect(self, x1, y1, x2, y2, radius=10, **kwargs):
        points = []
        corners = [
            (x1 + radius, y1 + radius, 3.14159, 4.71238),  # Oben links
            (x2 - radius, y1 + radius, 4.71238, 6.28318),  # Oben rechts
            (x2 - radius, y2 - radius, 0, 1.57079),        # Unten rechts
            (x1 + radius, y2 - radius, 1.57079, 3.14159)   # Unten links
        ]
        for cx, cy, start_angle, end_angle in corners:
            for i in range(9):  # 8 Segmente pro Ecke für Glätte
                angle = start_angle + (end_angle - start_angle) * i / 8
                points.extend([cx + radius * math.cos(angle),
                                cy + radius * math.sin(angle)])
        return self.create_polygon(points, smooth=True, **kwargs)

# ========== Funktionen ==========
def lade_beraterliste():
    """Lädt Beraterliste aus Excel-Datei"""
    global berater_df
    try:
        if os.path.exists(BERATER_LISTE):
            berater_df = pd.read_excel(BERATER_LISTE, dtype=str).fillna("")
            logging.info(f"Beraterliste geladen: {len(berater_df)} Einträge")
        else:
            logging.warning("Beraterliste nicht gefunden")
            berater_df = pd.DataFrame(columns=["Berater_Name", "Berater_Beraternummer"])
    except Exception as e:
        logging.error(f"Fehler beim Laden der Beraterliste: {e}")
        berater_df = pd.DataFrame(columns=["Berater_Name", "Berater_Beraternummer"])

def lade_vorlagen_liste():
    """Gibt alle .docx Dateien aus dem Vorlagen-Ordner zurück"""
    if not VORLAGEN_PATH.exists():
        return []
    return sorted([f.name for f in VORLAGEN_PATH.glob("*.docx")])

def on_berater_auswahl(event):
    """Event-Handler für Berater-Auswahl"""
    name = cb_berater.get()
    if name and not berater_df.empty:
        row = berater_df[berater_df['Berater_Name'] == name]
        if not row.empty:
            # Aktualisiere globales Dictionary
            werte_dict.update({
                'Berater_Name': row['Berater_Name'].iloc[0],
                'Berater_Beraternummer': row['Berater_Beraternummer'].iloc[0],
                'Berater_Titel': row.get('Berater_Titel', '').iloc[0] or "N/A",
                'Berater_E-Mail': row.get('Berater_E-Mail', '').iloc[0],
                'Berater_Telefonnummer': row.get('Berater_Telefonnummer', '').iloc[0]
            })
            
            # GUI-Felder aktualisieren
            entry_name.config(state='normal')
            entry_nr.config(state='normal')
            entry_name.delete(0, tk.END)
            entry_name.insert(0, werte_dict['Berater_Name'])
            entry_nr.delete(0, tk.END)
            entry_nr.insert(0, werte_dict['Berater_Beraternummer'])
            entry_name.config(state='readonly')
            entry_nr.config(state='readonly')

def aktualisiere_create_button():
    """Aktiviert/Deaktiviert den Erstellen-Button je nach Auswahl"""
    state = "normal" if excel_datei and bericht_datei else "disabled"
    btn_create.config(state=state)

def on_vorlage_auswahl(event):
    """Event-Handler für Word-Vorlage-Auswahl aus Dropdown"""
    global bericht_datei
    name = cb_vorlage.get()
    if name:
        bericht_datei = str(VORLAGEN_PATH / name)
        logging.info(f"Word-Vorlage gewählt: {name}")
    aktualisiere_create_button()

# ========== Datei-Handling ==========
def lade_excel():
    """Excel-Datei auswählen"""
    global excel_datei
    datei = filedialog.askopenfilename(
        title="Excel-Tags auswählen",
        filetypes=[("Excel Dateien", "*.xlsx *.xls")]
    )
    if datei:
        excel_datei = datei
        lbl_excel.config(text=os.path.basename(datei))
        aktualisiere_create_button()
        logging.info(f"Excel-Datei ausgewählt: {os.path.basename(datei)}")

def import_word():
    """Andere Word-Vorlage auswählen (außerhalb des Vorlagen-Ordners)"""
    global bericht_datei
    datei = filedialog.askopenfilename(
        title="Word-Vorlage auswählen",
        filetypes=[("Word Dateien", "*.docx")]
    )
    if datei:
        bericht_datei = datei
        name = os.path.basename(datei)
        cb_vorlage.set(name)
        logging.info(f"Eigene Word-Vorlage gewählt: {name}")
    aktualisiere_create_button()

def handle_drop(event):
    """Drag & Drop Event"""
    global excel_datei
    excel_datei = event.data.strip('{}')
    lbl_excel.config(text=os.path.basename(excel_datei))
    aktualisiere_create_button()

def entferne_nicht_passende_massnahmen_sdt(doc, werte):
    """
    Verarbeitet alle 'Anzahl_Maßnahmen_X' Content Controls:
    - Nicht passende (falsche Zahl) → vollständig gelöscht
    - Passender (richtige Zahl)     → Wrapper entfernt, Inhalt bleibt als normaler Text
    """
    anzahl_wert = werte.get('Anzahl_Maßnahmen', '')
    try:
        anzahl = int(str(anzahl_wert).strip())
    except (ValueError, TypeError):
        logging.warning(f"Anzahl_Maßnahmen hat ungültigen Wert: '{anzahl_wert}' – keine SDTs verändert")
        return

    zu_loeschen = []
    zu_unwrappen = []

    for sdt in doc.element.iter():
        if sdt.tag.endswith('sdt'):
            tag_el = sdt.find('.//w:tag', namespaces=WORD_NS)
            if tag_el is not None:
                key = tag_el.get(qn('w:val')) or ''
                if key.startswith('Anzahl_Maßnahmen_'):
                    suffix = key[len('Anzahl_Maßnahmen_'):]
                    try:
                        if int(suffix) == anzahl:
                            zu_unwrappen.append(sdt)
                        else:
                            zu_loeschen.append(sdt)
                    except ValueError:
                        pass

    # Nicht passende vollständig löschen
    for sdt in zu_loeschen:
        parent = sdt.getparent()
        if parent is not None:
            parent.remove(sdt)

    # Passenden unwrappen: Wrapper weg, Inhalt bleibt
    for sdt in zu_unwrappen:
        parent = sdt.getparent()
        if parent is not None:
            idx = list(parent).index(sdt)
            content = sdt.find('w:sdtContent', namespaces=WORD_NS)
            if content is not None:
                for i, child in enumerate(list(content)):
                    parent.insert(idx + i, child)
            parent.remove(sdt)

    logging.info(f"Anzahl_Maßnahmen={anzahl}: {len(zu_loeschen)} gelöscht, {len(zu_unwrappen)} unwrapped")


def ersetze_content_controls(doc_path, werte, output_path):
    """Content Controls in Word ersetzen mit Tag-Validation"""
    try:
        doc = Document(doc_path)
        fehlende_tags = []

        # Nicht passende Anzahl_Maßnahmen_X Controls vor der Ersetzung löschen
        entferne_nicht_passende_massnahmen_sdt(doc, werte)
        
        for sdt in doc.element.iter():
            if sdt.tag.endswith('sdt'):
                tag_el = sdt.find('.//w:tag', namespaces=WORD_NS)
                if tag_el is not None:
                    key = tag_el.get(qn('w:val'))

                    # Prüfen ob Wert fehlt oder leer ist
                    if key not in werte or not str(werte[key]).strip():
                        fehlende_tags.append(key)
                    
                    content = sdt.find('.//w:sdtContent', namespaces=WORD_NS)
                    if content is not None:
                        texts = content.findall('.//w:t', namespaces=WORD_NS)
                        if texts:
                            texts[0].text = str(werte.get(key, ""))
                            for text_node in texts[1:]:
                                text_node.text = ""
        
        doc.save(output_path)
        
        # Zeige fehlende Tags an
        if fehlende_tags:
            zeige_fehlende_tags(fehlende_tags)
        
        return True
    except Exception as e:
        messagebox.showerror("Fehler beim Ersetzen", str(e))
        return False

def zeige_fehlende_tags(fehlende_tags):
    """Zeigt Fenster mit nicht ersetzten Tags"""
    tags = sorted(set(tag for tag in fehlende_tags if tag))
    if not tags:
        return
        
    fehlende_window = tk.Toplevel(root)
    fehlende_window.title("Nicht gefüllte Platzhalter im Bericht")
    fehlende_window.geometry("450x350")
    fehlende_window.configure(bg=COLORS['background'])
    fehlende_window.attributes('-topmost', True)
    fehlende_window.lift()
    fehlende_window.focus_force()
    fehlende_window.grab_set()  # Modal: blockiert das Hauptfenster bis geschlossen
    
    tk.Label(fehlende_window, text="Nicht gefüllte Platzhalter im Bericht",
             bg=COLORS['background'], font=("Arial", 14, "bold")).pack(pady=(10,5))
    tk.Label(fehlende_window,
             text="Bitte prüfe diese Tags in deiner Vorlage oder in der Excel-Datei.",
             bg=COLORS['background'], font=("Arial", 10), wraplength=400, justify="center").pack(pady=(0,10))
    
    frame = tk.Frame(fehlende_window, bg=COLORS['background'])
    frame.pack(fill="both", expand=True, padx=10, pady=5)
    
    text_widget = tk.Text(frame, wrap="none", bg="#ffffff", font=FONTS['label'])
    vsb = ttk.Scrollbar(frame, orient="vertical", command=text_widget.yview)
    text_widget.configure(yscrollcommand=vsb.set)
    vsb.pack(side="right", fill="y")
    text_widget.pack(side="left", fill="both", expand=True)
    
    text_widget.insert("1.0", "\n".join(tags))
    text_widget.config(state="disabled")
    
    tk.Button(fehlende_window, text="Schließen", command=fehlende_window.destroy,
              bg=COLORS['primary'], fg="white", font=FONTS['button']).pack(pady=10)
    
    logging.warning(f"Nicht gefüllte Tags: {tags}")

def bericht_erstellen():
    """Hauptfunktion: Bericht erstellen"""
    if not excel_datei or not bericht_datei:
        messagebox.showwarning("Fehler", "Bitte Excel- und Word-Datei auswählen!")
        return
    
    try:
        # Excel laden
        df = pd.read_excel(excel_datei, sheet_name='Export NWG', dtype=str).fillna("")
        if 'Tags' not in df.columns or 'Werte' not in df.columns:
            raise ValueError("Excel muss 'Tags' und 'Werte' Spalten haben")
        
        # Daten zusammenführen
        werte_dict.update(dict(zip(df['Tags'], df['Werte'])))
        
        # Speicherpfad
        adresse = werte_dict.get('Gebäude_Adresse', '')
        adresse_clean = "".join(c if c not in r'\/:*?"<>|' else "_" for c in adresse).strip()
        default_name = f"Sanierungsfahrplan_{adresse_clean}" if adresse_clean else "Sanierungsfahrplan"

        save_path = filedialog.asksaveasfilename(
            defaultextension='.docx',
            filetypes=[('Word Dokumente', '*.docx')],
            title="Bericht speichern unter",
            initialfile=default_name
        )
        
        if save_path and ersetze_content_controls(bericht_datei, werte_dict, save_path):
            messagebox.showinfo("Erfolg", f"Bericht gespeichert:\n{save_path}")
            logging.info(f"Bericht erstellt: {save_path}")
            
    except Exception as e:
        messagebox.showerror("Fehler", str(e))
        logging.error(f"Fehler beim Erstellen: {e}")

def show_easter_egg(event=None):
    """Easter Egg - Doppelklick auf Logo"""
    egg_win = tk.Toplevel(root)
    egg_win.title("🐣 Easter Egg")
    egg_win.geometry("350x150")
    egg_win.configure(bg=COLORS['background'])
    
    tk.Label(
        egg_win,
        text="Liebe Grüße\n vom aller Echten\n Elia Salemi\n 🦖",
        bg=COLORS['background'],
        fg=COLORS['primary'],
        font=("Arial", 18, "bold")
    ).pack(expand=True, fill="both", padx=20, pady=20)

# ========== GUI Aufbau ==========
logging.info("NWG-Bericht Converter gestartet - Clean Version")
lade_beraterliste()
vorlagen_liste = lade_vorlagen_liste()
bericht_datei = str(VORLAGEN_PATH / vorlagen_liste[0]) if vorlagen_liste else None

# Hauptfenster
root = TkinterDnD.Tk()
root.title("NWG-Bericht Converter")
root.geometry("1100x560")
root.resizable(False, False)
root.configure(bg=COLORS['background'])

# Icon setzen
if os.path.exists(ICON_PATH):
    try:
        root.iconbitmap(ICON_PATH)
    except Exception:
        pass

# Header
header = tk.Frame(root, bg=COLORS['primary'], height=60)
header.pack(fill='x')
tk.Label(header, text="NWG-Bericht Converter", bg=COLORS['primary'], 
         fg="white", font=FONTS['header']).place(relx=0.5, rely=0.5, anchor="center")

# Hauptbereich
main = tk.Frame(root, bg=COLORS['background'])
main.pack(fill='both', expand=True, padx=20, pady=20)
main.columnconfigure((0,1,2), weight=1)
main.rowconfigure(0, weight=1)

# Linke Spalte: Energieberater
frm_left = tk.LabelFrame(main, text="Energieberater", bg=COLORS['background'], 
                         fg=COLORS['text'], font=('Arial',12,'bold'), padx=20, pady=20)
frm_left.grid(row=0, column=0, sticky='nsew', padx=10, pady=10)
frm_left.columnconfigure(1, weight=1)

tk.Label(frm_left, text="Auswahl:", bg=COLORS['background'], font=FONTS['label']).grid(row=0, column=0, sticky='w')
cb_berater = ttk.Combobox(frm_left, values=list(berater_df['Berater_Name']), state='readonly', width=25)
cb_berater.grid(row=0, column=1, padx=(10,0), pady=5, sticky="ew")
cb_berater.bind('<<ComboboxSelected>>', on_berater_auswahl)

tk.Label(frm_left, text="Name:", bg=COLORS['background'], font=FONTS['label']).grid(row=1, column=0, sticky='w', pady=5)
entry_name = ttk.Entry(frm_left, width=25, state='readonly')
entry_name.grid(row=1, column=1, padx=(10,0))

tk.Label(frm_left, text="Berater-Nr.:", bg=COLORS['background'], font=FONTS['label']).grid(row=2, column=0, sticky='w', pady=5)
entry_nr = ttk.Entry(frm_left, width=25, state='readonly')
entry_nr.grid(row=2, column=1, padx=(10,0))

# Mittlere Spalte: Logo & Button
frm_center = tk.Frame(main, bg=COLORS['background'])
frm_center.grid(row=0, column=1, sticky='nsew', padx=10, pady=10)

# Logo mit Easter Egg
if os.path.exists(LOGO_PATH):
    try:
        img = Image.open(LOGO_PATH).resize((120, 120), Image.Resampling.LANCZOS)
        logo_img = ImageTk.PhotoImage(img)
        logo_label = tk.Label(frm_center, image=logo_img, bg=COLORS['background'])
        logo_label.pack(pady=(50,20))
        logo_label.bind("<Double-Button-1>", show_easter_egg)  # Easter Egg!
        frm_center.logo_img = logo_img
    except Exception:
        pass

# Hauptbutton
btn_create = ModernButton(frm_center, "🚀 Bericht erstellen", bericht_erstellen, 
                        width=220, height=48, bg=COLORS['primary'])
btn_create.pack(pady=20)
btn_create.config(state="disabled")

# Rechte Spalte: Import
frm_right = tk.LabelFrame(main, text="Import", bg=COLORS['background'],
                         fg=COLORS['text'], font=('Arial',12,'bold'), padx=15, pady=10)
frm_right.grid(row=0, column=2, sticky='nsew', padx=10, pady=10)

# --- Excel-Bereich ---
tk.Label(frm_right, text="Pfadfinder:", bg=COLORS['background'],
         font=FONTS['label']).pack(anchor='w', pady=(4,0))

drop_frame = tk.Frame(frm_right, bg=COLORS['drop_bg'], height=90)
drop_frame.pack(fill='x', pady=(4,6))
drop_frame.pack_propagate(False)

drop_canvas = tk.Canvas(drop_frame, bg=COLORS['drop_bg'], highlightthickness=0, height=90)
drop_canvas.pack(fill='both', expand=True)

def draw_drop_zone(event):
    drop_canvas.delete("all")
    w, h = event.width, event.height
    drop_canvas.create_rectangle(5, 5, w-5, h-5, dash=(5,3), outline="#34C759", width=2)
    drop_canvas.create_text(w//2, h//2, text="Pfadfinder hier ablegen",
                           fill="#34C759", font=FONTS['label'])

drop_canvas.bind("<Configure>", draw_drop_zone)
drop_canvas.drop_target_register(DND_FILES)
drop_canvas.dnd_bind('<<Drop>>', handle_drop)

lbl_excel = tk.Label(frm_right, text="Keine Datei gewählt", font=FONTS['label'],
                     bg=COLORS['background'], fg=COLORS['text'])
lbl_excel.pack(fill='x', pady=(2,6))

ModernButton(frm_right, "➕ Pfadfinder auswählen", lade_excel, width=180, height=38,
           bg="#34C759").pack(pady=(0,14))

# --- Trennlinie ---
ttk.Separator(frm_right, orient='horizontal').pack(fill='x', pady=(0,14))

# --- Word-Vorlage-Bereich ---
tk.Label(frm_right, text="Word-Vorlage:", bg=COLORS['background'],
         font=FONTS['label']).pack(anchor='w')
cb_vorlage = ttk.Combobox(frm_right, values=vorlagen_liste, state='readonly', width=24)
cb_vorlage.pack(fill='x', pady=(4,8))
cb_vorlage.bind('<<ComboboxSelected>>', on_vorlage_auswahl)

if vorlagen_liste:
    cb_vorlage.set(vorlagen_liste[0])

tk.Label(frm_right, text="oder", bg=COLORS['background'],
         fg=COLORS['text'], font=FONTS['label']).pack(pady=(0,4))
ModernButton(frm_right, "Import Word-Bericht", import_word, width=180, height=38,
           bg="#007AFF", hover_bg="#0051D4").pack()

root.mainloop()
