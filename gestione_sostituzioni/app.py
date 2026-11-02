"""
gestione_sostituzioni.app
Classe App: finestra principale dell'applicazione Tkinter.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import re
import os
import sys
import traceback
import webbrowser
import tempfile
import threading
import logging
from datetime import datetime
from typing import Dict, List

from .config import Configuration
from .utils import to_slug, normalize_class
from .data_manager import DataManager
from .database import DatabaseHandler
from .engine import SubstitutionEngine
from .editor import EditorWindow
from .exporters import HAS_DOCX, HAS_PANDAS, set_docx_cell_shading

logger = logging.getLogger("GestioneSostituzioni.App")

# Import condizionali per esportazione
if HAS_DOCX:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH

if HAS_PANDAS:
    import pandas as pd


class App(tk.Tk):
    """Finestra principale dell'applicazione Gestione Sostituzioni."""

    def __init__(self):
        super().__init__()
        self.dm = DataManager()
        self.db = DatabaseHandler()
        self.db.setup()
        self.db.load_state(self.dm)
        self.engine = SubstitutionEngine(self.dm)

        self.title("Gestione Sostituzioni 3.3")
        self.geometry("1450x1000")
        self.configure(bg=Configuration.THEME['bg'])

        self.setup_icons()
        self.file_boxes = {}
        self.loaded_paths = {}  # Track file paths
        self.temp_files = []    # Lista file temporanei da pulire alla chiusura
        self.teacher_abbr_map = {}  # Mappatura etichetta tendina -> nominativo abbreviato
        self.setup_ui()
        self.setup_menu()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        logger.info("Applicazione avviata")

    def setup_icons(self):
        """Configura le icone dell'applicazione."""
        ico_path = Configuration.resource_path("Alberghiero.ico")
        png_path = Configuration.resource_path("alberghiero.png")
        if os.path.exists(png_path):
            try:
                self.iconphoto(True, tk.PhotoImage(file=png_path))
            except Exception:
                pass
        elif os.path.exists(ico_path):
            try:
                self.iconbitmap(ico_path)
            except Exception:
                pass
        self.logo_path = png_path

    def setup_menu(self):
        """Configura la barra dei menu."""
        menubar = tk.Menu(self)

        # File Menu
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Resetta tutto", command=self.reset_all_data)
        menubar.add_cascade(label="File", menu=file_menu)

        # Help Menu
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="📖 Manuale Utente (HTML)", command=self.open_manual_html)
        help_menu.add_command(label="Guida Rapida a schermo", command=self.show_guide)
        menubar.add_cascade(label="Aiuto", menu=help_menu)

        # About Menu
        about_menu = tk.Menu(menubar, tearoff=0)
        about_menu.add_command(label="About", command=self.show_about)
        menubar.add_cascade(label="About", menu=about_menu)

        self.config(menu=menubar)

    def open_manual_html(self):
        """Apre il Manuale Utente in formato HTML nel browser predefinito."""
        manual_path = Configuration.resource_path("manuale_utente.html")
        if not os.path.exists(manual_path):
            manual_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "manuale_utente.html")

        if os.path.exists(manual_path):
            try:
                webbrowser.open('file://' + os.path.abspath(manual_path))
            except Exception as e:
                logger.error(f"Errore apertura manuale: {e}")
                messagebox.showerror("Errore", f"Impossibile aprire il manuale:\n{e}")
        else:
            messagebox.showerror("Errore", "File manuale_utente.html non trovato.")

    def show_guide(self):
        """Mostra la guida rapida in una finestra popup."""
        top = tk.Toplevel(self)
        top.title("Guida Utilizzo")
        top.geometry("640x650")
        top.configure(bg="#FAFAFA")

        lbl = tk.Label(top, text="MANUALE UTENTE — GUIDA RAPIDA",
                       font=("Helvetica", 16, "bold"), bg="#FAFAFA", fg="#1565C0")
        lbl.pack(pady=(12, 4))

        btn_open_html = tk.Button(top, text="🌐 Apri Manuale Completo in HTML",
                                  font=("Segoe UI", 10, "bold"),
                                  bg="#E3F2FD", fg="#1565C0", relief="groove",
                                  padx=12, pady=4, cursor="hand2",
                                  command=self.open_manual_html)
        btn_open_html.pack(pady=(0, 8))

        txt_frame = tk.Frame(top, bg="#FAFAFA")
        txt_frame.pack(fill="both", expand=True, padx=15, pady=(0, 10))

        txt = tk.Text(txt_frame, wrap="word", bg="#FFFFFF", font=("Segoe UI", 10),
                      bd=1, relief="solid", padx=15, pady=10)
        vsb = ttk.Scrollbar(txt_frame, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        txt.pack(side="left", fill="both", expand=True)

        # Tags for styling
        txt.tag_config("h1", font=("Segoe UI", 12, "bold"), foreground="#1565C0", spacing3=6)
        txt.tag_config("h2", font=("Segoe UI", 11, "bold"), foreground="#2E7D32", spacing3=4)
        txt.tag_config("body", foreground="#333", spacing1=3)
        txt.tag_config("bullet", foreground="#444", lmargin1=15, lmargin2=15, spacing1=2)

        content = [
            ("1. Caricamento Dati (CSV)", "h1"),
            ("Caricare i file CSV necessari utilizzando i 6 riquadri colorati in alto. Cliccare su 'Carica' per selezionare il file.\nIl riquadro diventerà verde al caricamento. Cliccando su 'Anteprima' è possibile visualizzare, correggere e salvare i dati direttamente nel programma.", "body"),
            ("\n2. Inserimento Giornaliero Assenze", "h1"),
            ("Utilizzare i box di testo per inserire le informazioni della giornata odierna:", "body"),
            ("• Docenti Assenti: Inserire cognome e nome dei docenti assenti per malattia/permesso (incrementa le statistiche).", "bullet"),
            ("• Permessi Brevi: Inserire docente e ore (es. ROSSI 2 3 per 2ª e 3ª ora).", "bullet"),
            ("• Classi Assenti: Inserire le classi fuori sede (gite, uscite) per evitare sostituzioni inutili.", "bullet"),
            ("\n3. Elaborazione Sostituzioni", "h1"),
            ("Selezionare il giorno della settimana e cliccare sul pulsante verde 'PROCEDI'.\nLe ore da coprire derivano dall'Orario Generale e dall'Orario Personale. Il programma applica la gerarchia: Compresenza → Compresenza Sostegno (riquadro 6) → Disp. da classe assente → Potenziamento (round-robin) → Ore a pagamento → Classe scoperta.\nI nomi dei docenti devono essere scritti nello stesso modo in tutti i file e nelle caselle (solo cognome; per gli omonimi cognome+iniziale attaccati, es. BassuI).", "body"),
            ("\n4. Menu Contestuale (Tasto Destro)", "h1"),
            ("Facendo clic con il tasto destro su qualsiasi riga della tabella è possibile:\n- Segnare la riga come 'Coperta manualmente' (sfondo verde salvia);\n- Modificare manualmente il testo del sostituto;\n- Rimuovere la riga dalla tabella;\n- Copiare i dati negli appunti.", "body"),
            ("\n5. Esportazione e Stampa", "h1"),
            ("• PDF (Blu): Anteprima di stampa con layout ufficiale dell'Istituto.\n• XLSX (Verde): Esportazione foglio di calcolo Excel.\n• WORD (Rosso): Esportazione DOCX con scelta tra stile 'A Colori' (con evidenziazioni fedeli) o 'Bianco e Nero' per stampa classica.", "body"),
            ("\n⭐ Novità della Versione 3.3 (31 Ottobre 2026)", "h2"),
            ("• Tendine a comparsa (dropdown) integrate nei riquadri giornalieri: selezione rapida e senza errori di battitura per Docenti Assenti, Permessi Brevi e Classi Assenti.", "bullet"),
            ("• Autocompletamento e abbreviazioni automatiche coerenti con il database per evitare conflitti di omonimia (es. BASSUI/BASSUP, CANUL/CANUPA/CANUP).", "bullet"),
            ("• Selettore rapido delle ore per i Permessi Brevi (ore 1-8 con pulsante Aggiungi).", "bullet"),
            ("• Piena compatibilità con l'inserimento manuale/copia-incolla tradizionale mantenuto attivo.", "bullet"),
            ("\n⭐ Novità della Versione 3.1", "h2"),
            ("• Classi con due blocchi nell'orario generale (laboratorio cucina / sala) unite nello stesso orario.", "bullet"),
            ("• Docenti dei gruppi sdoppiati non proposti come compresenza l'uno dell'altro.", "bullet"),
            ("• Selezione file Excel corretta su macOS.", "bullet"),
            ("• Manuale con regole di scrittura dei nomi e priorità con Orario Personale e Sostegno.", "bullet"),
            ("\n⭐ Novità della Versione 3.0", "h2"),
            ("• Esportazione Word A Colori / B-N con selezione grafica personalizzata.", "bullet"),
            ("• Gestione Materia Alternativa corretta: sostituzione per gli alunni non avvalentesi quando manca solo l'alternativa (con religione presente), e singolo docente sostituto se mancano entrambi.", "bullet"),
            ("• Grafica tabella ad alta leggibilità (rowheight 52px e colonna espandibile): nessun testo tagliato.", "bullet"),
            ("• Avvio ripulito: le caselle giornaliere partono sempre pulite e pronte per il nuovo giorno.", "bullet"),
            ("• Algoritmo di potenziamento round-robin per un'equa rotazione dei docenti.", "bullet"),
            ("• Editor CSV anti-troncamento: salvataggio sicuro e zero perdita di righe.", "bullet"),
            ("• Menu contestuale tasto destro completo per gestione rapida delle righe.", "bullet"),
            ("• Architettura modulare con logging strutturato per manutenibilità e diagnostica.", "bullet"),
        ]

        for text, tag in content:
            txt.insert("end", text + "\n", tag)

        txt.config(state="disabled")

    def show_about(self):
        """Mostra la finestra About."""
        top = tk.Toplevel(self)
        top.title("About")
        top.geometry("420x330")
        top.configure(bg="white")
        top.resizable(False, False)

        tk.Label(top, text="Gestione Sostituzioni", font=("Helvetica", 22, "bold"),
                 bg="white", fg="#2c3e50").pack(pady=(30, 8))
        tk.Label(top, text="Versione 3.3", font=("Segoe UI", 13, "bold"),
                 bg="white", fg="#1976D2").pack(pady=2)
        tk.Label(top, text="Ultimo aggiornamento: 31 Ottobre 2026",
                 font=("Segoe UI", 11), bg="white", fg="#555555").pack(pady=4)
        tk.Label(top, text="Mario Carta copyright © 2026",
                 font=("Segoe UI", 11, "italic"), bg="white", fg="#2c3e50").pack(pady=(25, 15))
        tk.Button(top, text="Chiudi", font=("Segoe UI", 10), bg="#F5F5F5",
                  relief="groove", padx=15, pady=3, command=top.destroy).pack(pady=10)

    def reset_all_data(self):
        """Resetta completamente tutti i dati, statistiche e database."""
        if not messagebox.askyesno("Attenzione",
                                    "Questa operazione è irreversibile!\n\nTutti i dati caricati, le statistiche e il database verranno cancellati completamente.\n\nVuoi davvero procedere?"):
            return

        # 1. Clear Memory Data
        self.dm = DataManager()
        self.engine = SubstitutionEngine(self.dm)
        self.loaded_paths = {}

        # 2. Clear UI Inputs
        self.txt_absent.delete("1.0", "end")
        self.txt_short_leave.delete("1.0", "end")
        self.txt_classes_absent.delete("1.0", "end")
        self.entry_teacher.delete(0, "end")
        self.entry_total_teachers.delete(0, "end")
        self.entry_total_teachers.insert(0, "0")

        # 3. Clear Table
        self.clear_table()

        # 4. Wipe Database
        try:
            with self.db.get_connection() as conn:
                conn.execute("DELETE FROM daily_stats")
                conn.execute("DELETE FROM settings")
                conn.execute("DELETE FROM data_store")
                conn.execute("DELETE FROM preview_store")
                conn.commit()
        except Exception as e:
            logger.error(f"Errore pulizia database: {e}")

        # 5. Refresh UI
        self.refresh_labels()
        self.update_stats_display()

        logger.info("Reset completo eseguito")
        messagebox.showinfo("Reset Completato", "Tutti i dati sono stati cancellati e il programma è stato azzerato.")

    def setup_ui(self):
        """Configura l'interfaccia utente principale."""
        style = ttk.Style()
        style.theme_use('clam')
        bg = Configuration.THEME['bg']
        primary = Configuration.THEME['primary']

        style.configure("Treeview", background="#FFFFFF", foreground="black",
                         rowheight=52, font=("Segoe UI", 11), borderwidth=1)
        style.configure("Treeview.Heading", font=("Segoe UI", 12, "bold"),
                         background=primary, foreground="white", relief="flat")
        style.layout("Treeview", [('Treeview.treearea', {'sticky': 'nswe'})])

        # Header
        header_frame = tk.Frame(self, bg=primary)
        header_frame.pack(fill="x")
        h_inner = tk.Frame(header_frame, bg=primary)
        h_inner.pack(pady=5, padx=30, anchor="w")
        if self.logo_path and os.path.exists(self.logo_path):
            try:
                self.img_header = tk.PhotoImage(file=self.logo_path).subsample(6, 6)
                tk.Label(h_inner, image=self.img_header, bg=primary).pack(side="left", padx=(0, 15))
            except Exception:
                pass
        tk.Label(h_inner, text="GESTIONE SOSTITUZIONI",
                 font=("Helvetica", 24, "bold"), fg="white", bg=primary).pack(side="left")

        # Copyright on the right of the title
        tk.Label(h_inner, text="© Mario Carta", font=("Helvetica", 10),
                 fg="white", bg=primary).pack(side="left", padx=(10, 0), anchor="cen")

        # STATS FRAME
        stats_frame = tk.Frame(self, bg="#E3F2FD", bd=1, relief="solid", padx=20, pady=2)
        stats_frame.pack(fill="x", padx=30, pady=2)

        tk.Label(stats_frame, text="STATISTICHE ASSENZE",
                 font=("Segoe UI", 11, "bold"), bg="#E3F2FD", fg="#1565C0").pack(side="left")

        tk.Label(stats_frame, text="Totale Organico:", bg="#E3F2FD", fg="#000000",
                 font=("Segoe UI", 10)).pack(side="left", padx=(20, 5))
        self.entry_total_teachers = tk.Entry(stats_frame, width=5, bg="white", fg="black",
                                              insertbackground="black")
        self.entry_total_teachers.pack(side="left")
        self.entry_total_teachers.insert(0, str(self.dm.total_teachers))

        self.lbl_stats = tk.Label(stats_frame, text="Caricamento...", bg="#E3F2FD", fg="#000000",
                                   font=("Segoe UI", 10))
        self.lbl_stats.pack(side="left", padx=20)

        tk.Button(stats_frame, text="Resetta Statistiche", bg="#ffcdd2",
                  command=self.reset_stats, font=("Segoe UI", 8)).pack(side="right", padx=5)
        tk.Button(stats_frame, text="Stampa Statistiche", bg="#FFF",
                  command=self.print_stats, font=("Segoe UI", 8)).pack(side="right", padx=5)

        self.update_stats_display()

        main_frame = tk.Frame(self, bg=bg, padx=10, pady=2)
        main_frame.pack(fill="both", expand=True)

        # GRID FILES
        grid_frame = tk.Frame(main_frame, bg=bg)
        grid_frame.pack(fill="x", pady=2)
        self.create_file_box(grid_frame, "1. ORARIO CLASSI", "main", 0, 0, "#00695C",
                             lambda p: self.dm.parse_main(p))
        self.create_file_box(grid_frame, "2. ORARIO PERSONALE", "pers", 0, 1, "#0277BD",
                             lambda p: self.dm.parse_pers(p, 'pers'))
        self.create_file_box(grid_frame, "3. MATERIA ALTERNATIVA", "alt", 0, 2, "#6A1B9A",
                             lambda p: self.dm.parse_alt(p))
        self.create_file_box(grid_frame, "4. POTENZIAMENTO", "pot", 1, 0, "#EF6C00",
                             lambda p: self.dm.parse_generic(p, 'pot'))
        self.create_file_box(grid_frame, "5. A PAGAMENTO", "pay", 1, 1, "#2E7D32",
                             lambda p: self.dm.parse_generic(p, 'pay'))
        self.create_file_box(grid_frame, "6. DOCENTI SOSTEGNO", "sos", 1, 2, "#d35400",
                             lambda p: self.dm.parse_sos(p))

        # TEXT AREAS
        lists_frame = tk.Frame(main_frame, bg=bg)
        lists_frame.pack(fill="x", pady=2)
        lists_frame.columnconfigure(0, weight=1)
        lists_frame.columnconfigure(1, weight=1)
        lists_frame.columnconfigure(2, weight=1)

        self.txt_absent, self.cb_absent = self.create_input_box_with_dropdown(
            lists_frame, "DOCENTI ASSENTI", 0, 0, "#c0392b", bg_color="#FFEBEE", mode="teacher"
        )
        self.txt_short_leave, self.cb_short = self.create_input_box_with_dropdown(
            lists_frame, "PERMESSI BREVI", 0, 1, "#16a085", bg_color="#E0F7FA", mode="short_leave",
            hint="Es: Rossi 1 2"
        )
        self.txt_classes_absent, self.cb_classes = self.create_input_box_with_dropdown(
            lists_frame, "CLASSI ASSENTI", 0, 2, "#f39c12", bg_color="#FFFDE7", mode="class"
        )

        # CONTROLS
        ctrl_frame = tk.Frame(main_frame, bg=bg)
        ctrl_frame.pack(fill="x", pady=8)

        # Left side control group
        c_left = tk.Frame(ctrl_frame, bg=bg)
        c_left.pack(side="left", fill="x", expand=True)

        tk.Label(c_left, text="GIORNO:", font=("Segoe UI", 12, "bold"),
                 bg=bg, fg=Configuration.THEME['text']).pack(side="left")
        self.combo_day = ttk.Combobox(c_left, values=Configuration.DAYS[1:],
                                       state="readonly", width=12, font=("Segoe UI", 12))
        self.combo_day.current(0)
        self.combo_day.pack(side="left", padx=10)

        tk.Label(c_left, text="DOCENTI DA SOSTITUIRE:", font=("Segoe UI", 12, "bold"),
                 bg=bg, fg=Configuration.THEME['text']).pack(side="left", padx=(15, 5))
        self.entry_teacher = tk.Entry(c_left, font=("Segoe UI", 13), relief="flat", bd=2,
                                       highlightthickness=1, highlightbackground="#ccc",
                                       bg="white", fg="black", insertbackground="black")
        self.entry_teacher.pack(side="left", fill="x", expand=True, padx=5)

        # Right side buttons
        btn_frame = tk.Frame(ctrl_frame, bg=bg)
        btn_frame.pack(side="right", padx=5)
        self.create_btn(btn_frame, "PROCEDI", "#4CAF50", self.run_calculation, fg="white")
        self.create_btn(btn_frame, "SVUOTA", "#f44336", self.clear_table, fg="white")

        # Export Bar
        export_frame = tk.Frame(main_frame, bg=bg)
        export_frame.pack(fill="x", pady=2)
        self.create_btn(export_frame, "ANTEPRIMA STAMPA (PDF)", "#1565C0",
                         self.print_preview, side="left", fg="white")
        if HAS_PANDAS:
            self.create_btn(export_frame, "XLSX", "#2E7D32",
                             self.export_excel, side="left", fg="white")
        if HAS_DOCX:
            self.create_btn(export_frame, "WORD", "#c62828",
                             self.export_word, side="left", fg="white")

        self.create_btn(export_frame, "IMPORTA DB", "#FFFFFF",
                         self.import_db_file, side="right")
        self.create_btn(export_frame, "SALVA DB", "#FFFFFF",
                         self.export_db_file, side="right")

        tree_frame = tk.Frame(main_frame, bg="gray", bd=1, relief="solid")
        tree_frame.pack(fill="both", expand=True, pady=5)
        self.tree = ttk.Treeview(tree_frame, columns=("class", "hour", "orig", "sub"),
                                  show="headings", height=15)
        self.tree.heading("class", text="CLASSE")
        self.tree.column("class", width=90, minwidth=70, anchor="center")
        self.tree.heading("hour", text="ORA")
        self.tree.column("hour", width=60, minwidth=50, anchor="center")
        self.tree.heading("orig", text="SITUAZIONE")
        self.tree.column("orig", width=320, minwidth=250, anchor="w")
        self.tree.heading("sub", text="SOSTITUZIONE")
        self.tree.column("sub", width=650, minwidth=400, anchor="w", stretch=True)

        sb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=sb.set)
        sb.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)
        for k, col in Configuration.COLORS.items():
            self.tree.tag_configure(k, background=col)
        # Tag aggiuntivo per righe marcate come coperte
        self.tree.tag_configure('coperta', background='#A5D6A7')

        # Menu contestuale (tasto destro)
        self.setup_context_menu()

        self.refresh_labels()

    # --- MENU CONTESTUALE ---

    def setup_context_menu(self):
        """Crea e configura il menu contestuale per la tabella risultati."""
        self.ctx_menu = tk.Menu(self, tearoff=0)
        self.ctx_menu.add_command(label="✅ Segna come Coperta", command=self.mark_covered)
        self.ctx_menu.add_command(label="✏️ Modifica Sostituto", command=self.edit_substitute)
        self.ctx_menu.add_separator()
        self.ctx_menu.add_command(label="❌ Rimuovi riga", command=self.remove_row)
        self.ctx_menu.add_command(label="📋 Copia riga", command=self.copy_row)

        # Bind: tasto destro (Button-2 su Mac, Button-3 su Windows/Linux)
        if sys.platform == 'darwin':
            self.tree.bind("<Button-2>", self.show_context_menu)
            self.tree.bind("<Control-Button-1>", self.show_context_menu)
        else:
            self.tree.bind("<Button-3>", self.show_context_menu)

        # Doppio clic per modifica rapida sostituto
        self.tree.bind("<Double-1>", lambda e: self.edit_substitute())

    def show_context_menu(self, event):
        """Mostra il menu contestuale alla posizione del click."""
        row_id = self.tree.identify_row(event.y)
        if row_id:
            self.tree.selection_set(row_id)
            self.ctx_menu.tk_popup(event.x_root, event.y_root)

    def mark_covered(self):
        """Segna la riga selezionata come 'Coperta manualmente'."""
        selected = self.tree.selection()
        if not selected:
            return
        for item in selected:
            vals = list(self.tree.item(item, "values"))
            vals[3] = "COPERTA MANUALMENTE"
            self.tree.item(item, values=vals, tags=("coperta",))

    def edit_substitute(self):
        """Permette di selezionare o modificare il sostituto per la riga selezionata."""
        selected = self.tree.selection()
        if not selected:
            return
        item = selected[0]
        vals = list(self.tree.item(item, "values"))
        tags = list(self.tree.item(item, "tags"))
        if len(vals) < 4:
            return

        cls_val = str(vals[0])
        hour_val = str(vals[1])
        sit_val = str(vals[2])
        current_sub = str(vals[3])

        # Estrai ora come intero se possibile
        try:
            h_int = int(re.search(r'\d+', hour_val).group(0))
        except Exception:
            h_int = 1

        day = self.combo_day.get()

        # Raccoglie i candidati disponibili
        candidates = []

        # 1. Candidati presenti nel testo della sostituzione attuale
        # es: "DISP. DA CLASSE ASSENTE: BIANCHI (altri: SALARIS, NERI)"
        # oppure: "COMPRESENZA: ROSSI (o BIANCHI)"
        # oppure: "PAGAMENTO: VERDI (altri: ROSSI)"
        m_disp = re.match(r'^DISP\.\s*DA\s*CLASSE\s*ASSENTE:\s*([^\(]+)(?:\((?:altri|o):\s*([^\)]+)\))?', current_sub, re.IGNORECASE)
        if m_disp:
            main_t = m_disp.group(1).strip()
            if main_t:
                candidates.append(f"DISP. DA CLASSE ASSENTE: {main_t}")
            if m_disp.group(2):
                for extra in m_disp.group(2).split(','):
                    extra = extra.strip()
                    if extra:
                        candidates.append(f"DISP. DA CLASSE ASSENTE: {extra}")

        # 2. Tutti i docenti liberati da classi assenti per quell'ora (da engine)
        if hasattr(self, 'engine') and hasattr(self.engine, 'last_freed_teachers'):
            freed_list = self.engine.last_freed_teachers.get(h_int, [])
            for ft in freed_list:
                opt = f"DISP. DA CLASSE ASSENTE: {ft}"
                if opt not in candidates:
                    candidates.append(opt)

        # 3. Compresenze e docenti di sostegno presenti in orario per quella classe e ora
        if hasattr(self, 'dm') and self.dm:
            # Sostegno programmato per quella classe
            for s_slug, s_data in self.dm.db_sos.items():
                if isinstance(s_data, dict) and day in s_data and h_int in s_data[day]:
                    sos_cls = s_data[day][h_int]
                    if sos_cls == cls_val or (len(sos_cls) >= 3 and (sos_cls in cls_val or cls_val in sos_cls)):
                        s_name = s_data.get('_name', s_slug)
                        opt = f"COMPRESENZA SOSTEGNO: {s_name}"
                        if opt not in candidates:
                            candidates.append(opt)

            # Potenziamento per quell'ora
            pots = self.dm.db_pot.get(day, {}).get(h_int, [])
            for p in pots:
                opt = f"POTENZIAMENTO: {p}"
                if opt not in candidates:
                    candidates.append(opt)

            # Ore a pagamento per quell'ora
            pays = self.dm.db_pay.get(day, {}).get(h_int, [])
            for p in pays:
                opt = f"PAGAMENTO: {p}"
                if opt not in candidates:
                    candidates.append(opt)

        # 4. Opzioni standard
        candidates.append("COPERTA MANUALMENTE")
        candidates.append("CLASSE SCOPERTA")

        # Rimuovi duplicati preservando l'ordine
        seen = set()
        clean_candidates = []
        for c in candidates:
            if c not in seen:
                seen.add(c)
                clean_candidates.append(c)

        # Creazione della finestra di dialogo modale
        dialog = tk.Toplevel(self)
        dialog.title(f"Scegli Sostituto — Classe {cls_val} ({hour_val}ª Ora)")
        dialog.geometry("560x320")
        dialog.configure(bg="#F8FAFC")
        dialog.transient(self)
        dialog.grab_set()
        dialog.resizable(False, False)

        # Header info
        hdr_frame = tk.Frame(dialog, bg="#1E293B", pady=10, padx=16)
        hdr_frame.pack(fill="x")
        tk.Label(hdr_frame, text=f"Classe: {cls_val}   |   Ora: {hour_val}ª",
                 font=("Segoe UI", 12, "bold"), fg="#38BDF8", bg="#1E293B").pack(anchor="w")
        tk.Label(hdr_frame, text=f"Situazione: {sit_val}",
                 font=("Segoe UI", 10), fg="#CBD5E1", bg="#1E293B").pack(anchor="w", pady=(2, 0))

        # Pulsanti Azione in basso (ancorati sempre a fondo finestra)
        btn_frame = tk.Frame(dialog, bg="#F8FAFC", padx=20)
        btn_frame.pack(side="bottom", fill="x", pady=(0, 15))

        # Contenuto centrale espandibile
        content_frame = tk.Frame(dialog, bg="#F8FAFC", padx=20, pady=12)
        content_frame.pack(fill="both", expand=True)

        tk.Label(content_frame, text="Seleziona dalla lista dei docenti disponibili a quest'ora:\n(oppure digita un messaggio personalizzato, es: 'La classe entra alle 9:15')",
                 font=("Segoe UI", 10, "bold"), bg="#F8FAFC", fg="#334155", justify="left").pack(anchor="w", pady=(0, 8))

        cmb_sub = ttk.Combobox(content_frame, values=clean_candidates, font=("Segoe UI", 11), width=48)
        cmb_sub.set(current_sub)
        cmb_sub.pack(fill="x", pady=5)
        cmb_sub.focus()

        def apply_selection():
            new_val = cmb_sub.get().strip()
            if new_val:
                vals[3] = new_val
                # Determina il tag corrispondente al testo selezionato
                new_tag = tags[0] if tags else "scoperta"
                if "COPERTA MANUALMENTE" in new_val:
                    new_tag = "coperta"
                elif "DISP. DA CLASSE ASSENTE" in new_val:
                    new_tag = "disp_class"
                elif "SOSTEGNO" in new_val:
                    new_tag = "sostegno"
                elif "COMPRESENZA" in new_val:
                    new_tag = "compresenza"
                elif "POTENZIAMENTO" in new_val:
                    new_tag = "potenziamento"
                elif "PAGAMENTO" in new_val:
                    new_tag = "pagamento"
                elif "CLASSE SCOPERTA" in new_val:
                    new_tag = "scoperta"

                self.tree.item(item, values=vals, tags=(new_tag,))
                logger.info(f"Sostituto aggiornato per {cls_val} ora {hour_val}: {new_val}")
            dialog.destroy()

        def make_btn(parent, text, bg_color, command):
            lbl = tk.Label(parent, text=text, font=("Segoe UI", 10, "bold"),
                           bg=bg_color, fg="white", padx=16, pady=6, cursor="hand2", relief="raised", bd=1)
            lbl.pack(side="right", padx=5)
            lbl.bind("<Button-1>", lambda e: command())
            lbl.bind("<Enter>", lambda e: lbl.config(bg="#1E293B" if bg_color != "#1E293B" else "#0F172A"))
            lbl.bind("<Leave>", lambda e: lbl.config(bg=bg_color))
            return lbl

        make_btn(btn_frame, "Annulla", "#64748B", dialog.destroy)
        make_btn(btn_frame, "Applica Sostituto", "#0284C7", apply_selection)

        dialog.bind("<Return>", lambda e: apply_selection())
        dialog.bind("<Escape>", lambda e: dialog.destroy())
        self.wait_window(dialog)

    def remove_row(self):
        """Rimuove la riga selezionata dalla tabella."""
        selected = self.tree.selection()
        if not selected:
            return
        for item in selected:
            self.tree.delete(item)

    def copy_row(self):
        """Copia il contenuto della riga selezionata negli appunti."""
        selected = self.tree.selection()
        if not selected:
            return
        item = selected[0]
        vals = self.tree.item(item, "values")
        text = "\t".join(str(v) for v in vals)
        self.clipboard_clear()
        self.clipboard_append(text)

    def create_btn(self, parent, text, bg, cmd, side="right", fg="black"):
        """Crea un pulsante stilizzato compatibile con Mac."""
        if sys.platform == 'darwin':
            fg = "black"

        tk.Button(parent, text=text, bg=bg, activebackground="#E0E0E0", fg=fg,
                  highlightbackground=bg, font=("Segoe UI", 9, "bold"),
                  command=cmd, padx=10, pady=3, relief="raised", borderwidth=2,
                  cursor="hand2").pack(side=side, padx=3)

    def create_file_box(self, parent, title, key, r, c, color, load_callback):
        """Crea un riquadro per il caricamento di un file CSV."""
        f = tk.Frame(parent, bg="white", highlightbackground=color,
                     highlightthickness=1, padx=5, pady=2)
        f.grid(row=r, column=c, sticky="ew", padx=5, pady=2)
        parent.columnconfigure(c, weight=1)
        tk.Label(f, text=title, bg="white", font=("Segoe UI", 10, "bold"),
                 fg="#2c3e50").pack(anchor="w")
        lbl = tk.Label(f, text="Nessun dato", bg="white", fg="#888",
                       font=("Segoe UI", 9))
        lbl.pack(anchor="w", pady=0)

        def on_click():
            filetypes = [
                ("Tutti i formati supportati (CSV, Excel)",
                 ("*.csv", "*.xlsx", "*.xls", "*.xlsm", "*.xltx", "*.xlt")),
                ("File Excel (*.xlsx, *.xls, *.xlsm)",
                 ("*.xlsx", "*.xls", "*.xlsm", "*.xltx", "*.xlt")),
                ("File CSV (*.csv)", ("*.csv",)),
                ("Tutti i file (*.*)", ("*.*",))
            ]
            path = filedialog.askopenfilename(filetypes=filetypes)
            if path:
                n = load_callback(path)
                lbl.config(text=f"Caricati: {n}", fg="#2E7D32")
                self.loaded_paths[key] = path
                self.db.save_state(self.dm)
                self.refresh_labels()
                logger.info(f"File caricato [{key}]: {path} ({n} elementi)")

        def on_preview():
            self.show_editor(key, title)

        def on_clear():
            if messagebox.askyesno("Svuota Dati", f"Sei sicuro di voler eliminare i dati di '{title}'?"):
                self.dm.clear_category(key)
                self.db.clear_category(key)
                self.loaded_paths[key] = ""
                self.refresh_labels()
                logger.info(f"Dati svuotati per la categoria [{key}]")
                messagebox.showinfo("Fatto", f"I dati di '{title}' sono stati eliminati.")

        tk.Frame(f, height=1, bg="#eee").pack(fill="x", pady=2)
        btn_box = tk.Frame(f, bg="white")
        btn_box.pack(fill="x")
        tk.Button(btn_box, text="👁 Anteprima", bg="#FAFAFA", fg="#333",
                  relief="flat", command=on_preview, font=("Segoe UI", 9)).pack(side="left")
        tk.Button(btn_box, text="🗑️ Svuota", bg="#FAFAFA", fg="#c0392b",
                  relief="flat", command=on_clear, font=("Segoe UI", 9)).pack(side="left", padx=4)
        tk.Button(btn_box, text="📂 Carica", bg="#F5F5F5", fg="black",
                  relief="flat", command=on_click, font=("Segoe UI", 9)).pack(side="right")
        self.file_boxes[key] = {'label': lbl, 'loader': load_callback}
        return f

    def show_editor(self, key, title):
        """Apre l'editor CSV per il file specificato."""
        content = self.dm.preview_data.get(key, "")
        limit = 400
        if key == 'main':
            limit = 600
        elif key == 'pers':
            limit = 1000
        elif key == 'sos':
            limit = 700

        def save_callback(new_content):
            try:
                if key == 'main':
                    self.dm.parse_main(new_content, is_path=False)
                elif key == 'pers':
                    self.dm.parse_pers(new_content, key='pers', is_path=False)
                elif key == 'sos':
                    self.dm.parse_sos(new_content, is_path=False)
                elif key == 'alt':
                    self.dm.parse_alt(new_content, is_path=False)
                elif key in ('pot', 'pay'):
                    self.dm.parse_generic(new_content, key=key, is_path=False)

                # Save to disk (solo per file CSV, evitando di corrompere file Excel binari)
                if key in self.loaded_paths and self.loaded_paths[key]:
                    disk_path = self.loaded_paths[key]
                    if disk_path.lower().endswith('.csv'):
                        try:
                            with open(disk_path, 'w', encoding='utf-8') as f:
                                f.write(new_content)
                        except Exception as e:
                            logger.error(f"Errore salvataggio su disco: {e}")

                self.db.save_state(self.dm)
                self.refresh_labels()
                messagebox.showinfo("Fatto", "Modifiche salvate con successo.")
            except Exception as e:
                logger.error(f"Errore salvataggio editor: {e}", exc_info=True)
                messagebox.showerror("Errore Salvataggio", str(e))

        EditorWindow(self, title, content, save_callback, max_rows=limit)

    def refresh_labels(self):
        """Aggiorna le etichette dei riquadri file con lo stato attuale."""
        if 'main' in self.file_boxes:
            count = len(self.dm.db_main)
            txt = f"Caricati: {count} classi" if count > 0 else "Nessun dato"
            self.file_boxes['main']['label'].config(text=txt, fg="#2E7D32" if count else "#888")

        if 'pers' in self.file_boxes:
            c = sum(1 for v in self.dm.db_pers.values() if isinstance(v, dict) and any(k != '_name' for k in v))
            tot_slots = sum(sum(len(v[d]) for d in v if d != '_name') for v in self.dm.db_pers.values() if isinstance(v, dict))
            self.file_boxes['pers']['label'].config(
                text=f"Docenti: {c} ({tot_slots} ore)" if tot_slots else ("Docenti: {c}" if c else "Nessun dato"),
                fg="#2E7D32" if (c or tot_slots) else "#888")

        if 'sos' in self.file_boxes:
            c = sum(1 for v in self.dm.db_sos.values() if isinstance(v, dict) and any(k != '_name' for k in v))
            tot_slots = sum(sum(len(v[d]) for d in v if d != '_name') for v in self.dm.db_sos.values() if isinstance(v, dict))
            self.file_boxes['sos']['label'].config(
                text=f"Docenti: {c} ({tot_slots} ore)" if tot_slots else "Nessun dato",
                fg="#2E7D32" if tot_slots else "#888")

        for k in ['alt', 'pot', 'pay']:
            if k in self.file_boxes:
                has_data = False
                if k == 'alt':
                    has_data = any(self.dm.db_alt[d] for d in self.dm.db_alt)
                elif k == 'pot':
                    has_data = any(self.dm.db_pot[d] for d in self.dm.db_pot)
                elif k == 'pay':
                    has_data = any(self.dm.db_pay[d] for d in self.dm.db_pay)
                self.file_boxes[k]['label'].config(
                    text="Dati caricati" if has_data else "Nessun dato",
                    fg="#2E7D32" if has_data else "#888")

        # Aggiorna anche le liste delle tendine di inserimento
        self.update_dropdown_lists()

    def get_selectable_teachers(self):
        """Genera l'elenco dei docenti per i menu a tendina con il relativo nominativo abbreviato."""
        teachers = []
        for slug, d in self.dm.db_pers.items():
            n = d.get('_name', slug).strip()
            if not n:
                continue
            abbr = n.replace(' ', '')
            if n in ('ARA F', 'ARA S', 'BASSU I', 'BASSU P'):
                abbr = n.replace(' ', '')
            elif n == 'CANU LAURA':
                abbr = 'CANUL'
            elif n == 'CANU PAOLA':
                abbr = 'CANUPA'
            elif n == 'CANU PIETRO':
                abbr = 'CANUP'
            elif n == 'PINNA GIOVANNI':
                abbr = 'PINNAGIO'
            elif n == 'PINNA GIOVANNICO':
                abbr = 'PINNAGIOVANNICO'
            elif n == 'SANNA A':
                abbr = 'SANNAA'
            elif n == 'SANNA GIOVANNA':
                abbr = 'SANNAGIO'
            elif n == 'SINI FRANCESCA':
                abbr = 'SINIFRA'
            elif n == 'IDINI MICHELA':
                abbr = 'IDINIMIC'
            elif n == 'MURA MARCELLO':
                abbr = 'MURAMAR'
            elif n == 'SERRA STEFANO':
                abbr = 'SERRA'
            elif n == 'SIAS CLAUDIA':
                abbr = 'SIAS'
            elif n == 'MACCIOCCU FABIO':
                abbr = 'MACCIOCCU'
            elif len(n.split()) == 1:
                abbr = n
            teachers.append((abbr, n, 'Curricolare'))

        for slug, d in self.dm.db_sos.items():
            n = d.get('_name', slug).strip()
            if not n or n in ('(provvisorio)', 'ASSENTE'):
                continue
            parts = n.split()
            sn = parts[0]
            if n == 'CARTA AGATA':
                abbr = 'CARTAA'
            elif n == 'CARTA MARIO':
                abbr = 'CARTAM'
            elif n == 'CUCCURU NICOLETTA':
                abbr = 'CUCCURUN'
            elif n == 'DORE GRAZIANO':
                abbr = 'DOREG'
            elif n == 'DORE PIERPAOLO':
                abbr = 'DOREPP'
            elif n == 'MELIS GRAZIANO':
                abbr = 'MELISG'
            elif n == 'MELONI FRANCESCA':
                abbr = 'MELONIF'
            elif n == 'MELONI GIOVANNA':
                abbr = 'MELONIG'
            elif n == 'MULAS MARIA':
                abbr = 'MULASM'
            elif n == 'MULAS GIULIANA':
                abbr = 'MULASGIU'
            elif n == 'PINNA GIAMPIERO':
                abbr = 'PINNAGIAMPIERO'
            elif n == 'PINNA GIOVANNI MARIA':
                abbr = 'PINNAGIO'
            elif n == 'PINTUS ANNA LINA':
                abbr = 'PINTUSA'
            elif n == 'TANCA GIANMICHELE':
                abbr = 'TANCAGM'
            else:
                abbr = sn
            teachers.append((abbr, n, 'Sostegno'))

        seen = set()
        display_list = []
        abbr_map = {}
        for abbr, full, role in sorted(teachers, key=lambda x: (x[0], x[1])):
            disp = f"{abbr} ({full})" if abbr != full else abbr
            if disp not in seen:
                seen.add(disp)
                display_list.append(disp)
                abbr_map[disp] = abbr
        return display_list, abbr_map

    def get_selectable_classes(self):
        """Genera l'elenco ordinato delle classi attive per il menu a tendina."""
        return sorted(list(self.dm.db_main.keys()))

    def update_dropdown_lists(self):
        """Aggiorna i valori dei menu a tendina quando vengono caricati o modificati orari."""
        teacher_list, self.teacher_abbr_map = self.get_selectable_teachers()
        classes_list = self.get_selectable_classes()

        if hasattr(self, 'cb_absent') and self.cb_absent:
            self.cb_absent['values'] = teacher_list
        if hasattr(self, 'cb_short') and self.cb_short:
            self.cb_short['values'] = teacher_list
        if hasattr(self, 'cb_classes') and self.cb_classes:
            self.cb_classes['values'] = classes_list

    def create_input_box_with_dropdown(self, parent, title, r, c, color, bg_color="white", mode="teacher", hint=None):
        """Crea un riquadro per l'immissione giornaliera completo di menu a tendina a comparsa."""
        f = tk.Frame(parent, bg=bg_color, highlightbackground=color,
                     highlightthickness=1, padx=5, pady=2)
        f.grid(row=r, column=c, sticky="nsew", padx=5, pady=2)
        parent.columnconfigure(c, weight=1)

        # Header con titolo
        h_frame = tk.Frame(f, bg=bg_color)
        h_frame.pack(fill="x", pady=(0, 2))
        tk.Label(h_frame, text=title, bg=bg_color, font=("Segoe UI", 10, "bold"),
                 fg="#2c3e50").pack(side="left")

        # Casella di testo principale (mantiene la scrittura e il copia-incolla libero)
        txt = tk.Text(f, height=5, font=("Arial", 14, "bold"), bd=1, relief="solid",
                      highlightthickness=0, bg=bg_color, fg="black", insertbackground="black")
        txt.pack(fill="both", expand=True, padx=2, pady=(2, 4))

        # Barra di immissione con tendina a comparsa
        dropdown_bar = tk.Frame(f, bg=bg_color)
        dropdown_bar.pack(fill="x", padx=2, pady=(2, 2))

        cb = None

        if mode == "teacher":
            tk.Label(dropdown_bar, text="Docente:", bg=bg_color,
                     font=("Segoe UI", 9, "bold"), fg="#333").pack(side="left", padx=(0, 3))
            cb = ttk.Combobox(dropdown_bar, state="normal", font=("Segoe UI", 9))
            cb.pack(side="left", fill="x", expand=True, padx=2)

            def add_absent_teacher():
                sel = cb.get().strip()
                if not sel:
                    return
                abbr = self.teacher_abbr_map.get(sel, sel.split()[0].upper())
                curr_txt = txt.get("1.0", "end").strip()
                existing = [x.strip() for x in curr_txt.splitlines() if x.strip()]
                if abbr not in existing:
                    existing.append(abbr)
                    txt.delete("1.0", "end")
                    txt.insert("1.0", "\n".join(existing) + "\n")
                cb.set("")

            def apply_absent_to_targets():
                """Apre un popup con checkbox per ciascun docente assente e applica i selezionati a 'Docenti da Sostituire'."""
                curr_txt = txt.get("1.0", "end").strip()
                teachers = [x.strip() for x in re.split(r'[\n,]+', curr_txt) if len(x.strip()) > 1]
                if not teachers:
                    return messagebox.showwarning("Attenzione", "Nessun docente presente nell'elenco degli assenti!")

                # Rimuovi duplicati preservando l'ordine
                seen = set()
                unique_teachers = []
                for t in teachers:
                    if t not in seen:
                        seen.add(t)
                        unique_teachers.append(t)

                popup = tk.Toplevel(self)
                popup.title("Seleziona Assenti da Sostituire")
                popup.geometry("380x420")
                popup.configure(bg="#F8FAFC")
                popup.transient(self)
                popup.grab_set()

                tk.Label(popup, text="Seleziona i docenti da inviare in sostituzione:",
                         font=("Segoe UI", 10, "bold"), bg="#F8FAFC", fg="#1E293B").pack(anchor="w", padx=12, pady=(10, 5))

                # Scrollable frame per i checkbutton
                container = tk.Frame(popup, bg="#CBD5E1", bd=1, relief="solid")
                container.pack(fill="both", expand=True, padx=12, pady=5)

                canvas = tk.Canvas(container, bg="white", highlightthickness=0)
                scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
                scroll_frame = tk.Frame(canvas, bg="white")

                scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
                canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
                canvas.configure(yscrollcommand=scrollbar.set)

                scrollbar.pack(side="right", fill="y")
                canvas.pack(side="left", fill="both", expand=True)

                vars_map = {}
                for t in unique_teachers:
                    var = tk.BooleanVar(value=True)
                    vars_map[t] = var
                    ck = tk.Checkbutton(scroll_frame, text=f"  {t}", variable=var,
                                        font=("Segoe UI", 11), bg="white", anchor="w",
                                        activebackground="white")
                    ck.pack(fill="x", padx=10, pady=2)

                btn_bar = tk.Frame(popup, bg="#F8FAFC")
                btn_bar.pack(fill="x", padx=12, pady=(5, 10))

                def toggle_all(val):
                    for v in vars_map.values():
                        v.set(val)

                def confirm():
                    selected = [t for t, v in vars_map.items() if v.get()]
                    if selected:
                        existing = [x.strip() for x in re.split(r'[,;]+', self.entry_teacher.get()) if x.strip()]
                        for s in selected:
                            if s not in existing:
                                existing.append(s)
                        self.entry_teacher.delete(0, "end")
                        self.entry_teacher.insert(0, ", ".join(existing))
                    popup.destroy()

                # Pulsanti stilizzati con tk.Label per compatibilità cromatica macOS
                btn_all = tk.Label(btn_bar, text=" Tutti ", font=("Segoe UI", 9, "bold"),
                                   bg="#FFFFFF", fg="#0F172A", relief="solid", bd=1,
                                   padx=8, pady=3, cursor="hand2")
                btn_all.pack(side="left", padx=(0, 4))
                btn_all.bind("<Button-1>", lambda e: toggle_all(True))
                btn_all.bind("<Enter>", lambda e: btn_all.configure(bg="#F1F5F9"))
                btn_all.bind("<Leave>", lambda e: btn_all.configure(bg="#FFFFFF"))

                btn_none = tk.Label(btn_bar, text=" Nessuno ", font=("Segoe UI", 9, "bold"),
                                    bg="#FFFFFF", fg="#0F172A", relief="solid", bd=1,
                                    padx=8, pady=3, cursor="hand2")
                btn_none.pack(side="left")
                btn_none.bind("<Button-1>", lambda e: toggle_all(False))
                btn_none.bind("<Enter>", lambda e: btn_none.configure(bg="#F1F5F9"))
                btn_none.bind("<Leave>", lambda e: btn_none.configure(bg="#FFFFFF"))

                btn_confirm = tk.Label(btn_bar, text=" ✓ Conferma e Applica ", font=("Segoe UI", 9, "bold"),
                                       bg="#16A34A", fg="#FFFFFF", relief="solid", bd=1,
                                       padx=12, pady=4, cursor="hand2")
                btn_confirm.pack(side="right")
                btn_confirm.bind("<Button-1>", lambda e: confirm())
                btn_confirm.bind("<Enter>", lambda e: btn_confirm.configure(bg="#15803D"))
                btn_confirm.bind("<Leave>", lambda e: btn_confirm.configure(bg="#16A34A"))

            # Pulsante Applica nel riquadro Docenti Assenti (Label button per garantire colore e testo leggibile su macOS)
            btn_apply = tk.Label(h_frame, text="⚡ Applica a 'Da Sostituire'", font=("Segoe UI", 8, "bold"),
                                 bg="#16A34A", fg="#FFFFFF", relief="solid", bd=1,
                                 padx=6, pady=2, cursor="hand2")
            btn_apply.pack(side="right")
            btn_apply.bind("<Button-1>", lambda e: apply_absent_to_targets())
            btn_apply.bind("<Enter>", lambda e: btn_apply.configure(bg="#15803D"))
            btn_apply.bind("<Leave>", lambda e: btn_apply.configure(bg="#16A34A"))

            btn_add = tk.Button(dropdown_bar, text="➕ Aggiungi", font=("Segoe UI", 8, "bold"),
                                bg="#E0E0E0", fg="#1E293B", relief="flat", padx=6, pady=1,
                                command=add_absent_teacher, cursor="hand2")
            btn_add.pack(side="right", padx=(2, 0))
            cb.bind("<Return>", lambda e: add_absent_teacher())

        elif mode == "short_leave":
            # Pulsante Applica per Permessi Brevi
            def apply_short_to_targets():
                """Apre un popup con checkbox per ciascun permesso breve e applica i docenti selezionati a 'Docenti da Sostituire'."""
                curr_txt = txt.get("1.0", "end").strip()
                permessi_entries = []
                for line in curr_txt.splitlines():
                    for entry in line.strip().split(','):
                        parts = entry.strip().split()
                        if len(parts) >= 2 and any(p.isdigit() for p in parts[1:]):
                            teacher = parts[0].upper()
                            hours = [p for p in parts[1:] if p.isdigit()]
                            permessi_entries.append((teacher, " ".join(hours)))

                if not permessi_entries:
                    return messagebox.showwarning("Attenzione", "Nessun permesso breve valido inserito!")

                popup = tk.Toplevel(self)
                popup.title("Seleziona Permessi da Sostituire")
                popup.geometry("380x420")
                popup.configure(bg="#F8FAFC")
                popup.transient(self)
                popup.grab_set()

                tk.Label(popup, text="Seleziona i docenti con permesso da sostituire:",
                         font=("Segoe UI", 10, "bold"), bg="#F8FAFC", fg="#1E293B").pack(anchor="w", padx=12, pady=(10, 5))

                container = tk.Frame(popup, bg="#CBD5E1", bd=1, relief="solid")
                container.pack(fill="both", expand=True, padx=12, pady=5)

                canvas = tk.Canvas(container, bg="white", highlightthickness=0)
                scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
                scroll_frame = tk.Frame(canvas, bg="white")

                scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
                canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
                canvas.configure(yscrollcommand=scrollbar.set)

                scrollbar.pack(side="right", fill="y")
                canvas.pack(side="left", fill="both", expand=True)

                vars_map = {}
                for idx, (t, hrs) in enumerate(permessi_entries):
                    key = f"{t}_{idx}"
                    var = tk.BooleanVar(value=True)
                    vars_map[key] = (t, var)
                    label_txt = f"  {t} (ore: {hrs})"
                    ck = tk.Checkbutton(scroll_frame, text=label_txt, variable=var,
                                        font=("Segoe UI", 11), bg="white", anchor="w",
                                        activebackground="white")
                    ck.pack(fill="x", padx=10, pady=2)

                btn_bar = tk.Frame(popup, bg="#F8FAFC")
                btn_bar.pack(fill="x", padx=12, pady=(5, 10))

                def toggle_all(val):
                    for _, var in vars_map.values():
                        var.set(val)

                def confirm():
                    selected = [t for _, (t, var) in vars_map.items() if var.get()]
                    if selected:
                        existing = [x.strip() for x in re.split(r'[,;]+', self.entry_teacher.get()) if x.strip()]
                        for s in selected:
                            if s not in existing:
                                existing.append(s)
                        self.entry_teacher.delete(0, "end")
                        self.entry_teacher.insert(0, ", ".join(existing))
                    popup.destroy()

                # Pulsanti stilizzati con tk.Label per compatibilità cromatica macOS
                btn_all = tk.Label(btn_bar, text=" Tutti ", font=("Segoe UI", 9, "bold"),
                                   bg="#FFFFFF", fg="#0F172A", relief="solid", bd=1,
                                   padx=8, pady=3, cursor="hand2")
                btn_all.pack(side="left", padx=(0, 4))
                btn_all.bind("<Button-1>", lambda e: toggle_all(True))
                btn_all.bind("<Enter>", lambda e: btn_all.configure(bg="#F1F5F9"))
                btn_all.bind("<Leave>", lambda e: btn_all.configure(bg="#FFFFFF"))

                btn_none = tk.Label(btn_bar, text=" Nessuno ", font=("Segoe UI", 9, "bold"),
                                    bg="#FFFFFF", fg="#0F172A", relief="solid", bd=1,
                                    padx=8, pady=3, cursor="hand2")
                btn_none.pack(side="left")
                btn_none.bind("<Button-1>", lambda e: toggle_all(False))
                btn_none.bind("<Enter>", lambda e: btn_none.configure(bg="#F1F5F9"))
                btn_none.bind("<Leave>", lambda e: btn_none.configure(bg="#FFFFFF"))

                btn_confirm = tk.Label(btn_bar, text=" ✓ Conferma e Applica ", font=("Segoe UI", 9, "bold"),
                                       bg="#16A34A", fg="#FFFFFF", relief="solid", bd=1,
                                       padx=12, pady=4, cursor="hand2")
                btn_confirm.pack(side="right")
                btn_confirm.bind("<Button-1>", lambda e: confirm())
                btn_confirm.bind("<Enter>", lambda e: btn_confirm.configure(bg="#15803D"))
                btn_confirm.bind("<Leave>", lambda e: btn_confirm.configure(bg="#16A34A"))

            # Pulsante Applica per Permessi Brevi (Label button per garantire colore e testo leggibile su macOS)
            btn_apply_short = tk.Label(h_frame, text="⚡ Applica a 'Da Sostituire'", font=("Segoe UI", 8, "bold"),
                                       bg="#16A34A", fg="#FFFFFF", relief="solid", bd=1,
                                       padx=6, pady=2, cursor="hand2")
            btn_apply_short.pack(side="right")
            btn_apply_short.bind("<Button-1>", lambda e: apply_short_to_targets())
            btn_apply_short.bind("<Enter>", lambda e: btn_apply_short.configure(bg="#15803D"))
            btn_apply_short.bind("<Leave>", lambda e: btn_apply_short.configure(bg="#16A34A"))

            # Tendina docente
            tk.Label(dropdown_bar, text="Docente:", bg=bg_color,
                     font=("Segoe UI", 9, "bold"), fg="#333").pack(side="left", padx=(0, 2))
            cb = ttk.Combobox(dropdown_bar, state="normal", width=14, font=("Segoe UI", 9))
            cb.pack(side="left", fill="x", expand=True, padx=2)

            # Sottobarra per scelta ore
            hours_bar = tk.Frame(f, bg=bg_color)
            hours_bar.pack(fill="x", padx=2, pady=(1, 2))

            tk.Label(hours_bar, text="Ore:", bg=bg_color, font=("Segoe UI", 8, "bold"),
                     fg="#333").pack(side="left", padx=(0, 3))

            cb_from = ttk.Combobox(hours_bar, values=[f"{h}ª" for h in range(1, 9)],
                                   width=4, state="readonly", font=("Segoe UI", 8))
            cb_from.current(0)
            cb_from.pack(side="left", padx=1)

            tk.Label(hours_bar, text="a", bg=bg_color, font=("Segoe UI", 8),
                     fg="#666").pack(side="left", padx=1)

            cb_to = ttk.Combobox(hours_bar, values=[f"{h}ª" for h in range(1, 9)],
                                 width=4, state="readonly", font=("Segoe UI", 8))
            cb_to.current(0)
            cb_to.pack(side="left", padx=1)

            def add_short_leave():
                sel = cb.get().strip()
                if not sel:
                    return
                abbr = self.teacher_abbr_map.get(sel, sel.split()[0].upper())
                try:
                    h_start = int(cb_from.get().replace("ª", ""))
                    h_end = int(cb_to.get().replace("ª", ""))
                    if h_start > h_end:
                        h_start, h_end = h_end, h_start
                    hours_str = " ".join(str(h) for h in range(h_start, h_end + 1))
                except Exception:
                    hours_str = "1"

                line = f"{abbr} {hours_str}"
                curr_txt = txt.get("1.0", "end").strip()
                existing = [x.strip() for x in curr_txt.splitlines() if x.strip()]
                existing.append(line)
                txt.delete("1.0", "end")
                txt.insert("1.0", "\n".join(existing) + "\n")
                cb.set("")

            btn_add = tk.Button(hours_bar, text="➕ Aggiungi Permesso", font=("Segoe UI", 8, "bold"),
                                bg="#E0E0E0", fg="#1E293B", relief="flat", padx=6, pady=1,
                                command=add_short_leave, cursor="hand2")
            btn_add.pack(side="right", padx=(2, 0))
            cb.bind("<Return>", lambda e: add_short_leave())

        elif mode == "class":
            tk.Label(dropdown_bar, text="Classe:", bg=bg_color,
                     font=("Segoe UI", 9, "bold"), fg="#333").pack(side="left", padx=(0, 3))
            cb = ttk.Combobox(dropdown_bar, state="normal", font=("Segoe UI", 9))
            cb.pack(side="left", fill="x", expand=True, padx=2)

            def add_absent_class():
                sel = cb.get().strip()
                if not sel:
                    return
                curr_txt = txt.get("1.0", "end").strip()
                existing = [x.strip() for x in curr_txt.splitlines() if x.strip()]
                if sel not in existing:
                    existing.append(sel)
                    txt.delete("1.0", "end")
                    txt.insert("1.0", "\n".join(existing) + "\n")
                cb.set("")

            btn_add = tk.Button(dropdown_bar, text="➕ Aggiungi", font=("Segoe UI", 8, "bold"),
                                bg="#E0E0E0", fg="#1E293B", relief="flat", padx=6, pady=1,
                                command=add_absent_class, cursor="hand2")
            btn_add.pack(side="right", padx=(2, 0))
            cb.bind("<Return>", lambda e: add_absent_class())

        if hint:
            tk.Label(f, text=hint, bg=bg_color, fg="#7f8c8d",
                     font=("Segoe UI", 8)).pack(anchor="w", padx=2, pady=(1, 0))

        return txt, cb

    def create_styled_text_box(self, parent, title, r, c, color, hint=None, bg_color="white"):
        """Metodo di retrocompatibilità per creare caselle di testo semplici."""
        txt, _ = self.create_input_box_with_dropdown(parent, title, r, c, color, bg_color=bg_color, hint=hint)
        return txt

    def on_closing(self):
        """Gestione chiusura applicazione: pulizia file temp e salvataggio stato."""
        # Pulizia file temporanei HTML generati durante la sessione
        for temp_path in self.temp_files:
            try:
                os.unlink(temp_path)
            except Exception:
                pass

        # Le caselle giornaliere non vengono memorizzate tra sessioni
        self.dm.text_absent = ""
        self.dm.text_short = ""
        self.dm.text_classes = ""
        try:
            self.dm.total_teachers = int(self.entry_total_teachers.get())
        except Exception:
            pass
        self.db.save_state(self.dm)
        logger.info("Applicazione chiusa")
        self.destroy()

    # --- STATS LOGIC ---

    def update_stats_display(self):
        """Aggiorna la visualizzazione delle statistiche."""
        w, m, t, d, det = self.db.get_stats()
        self.lbl_stats.config(text=f"Settimanale: {w}%  |  Mensile: {m}%  |  Totale: {t}%")

    def reset_stats(self):
        """Azzera le statistiche con conferma."""
        if messagebox.askyesno("Conferma", "Vuoi davvero azzerare tutte le statistiche?"):
            self.db.reset_stats()
            self.update_stats_display()

    def print_stats(self):
        """Genera e apre nel browser un report HTML delle statistiche."""
        w, m, t, d, agg_details = self.db.get_stats()

        rows_html = ""
        sorted_dets = sorted(agg_details.items(), key=lambda x: x[1], reverse=True)
        for name, hours in sorted_dets:
            rows_html += f"<tr><td>{name}</td><td>{hours}</td></tr>"

        html = f"""<html><head><style>
            body {{ font-family: sans-serif; padding: 20px; }}
            .header {{ text-align: center; border-bottom: 2px solid #333; margin-bottom: 20px; }}
            .stat-box {{ border: 1px solid #ccc; padding: 10px; margin: 10px 0; background-color: #f9f9f9; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
            th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
            th {{ background-color: #f2f2f2; }}
            h1 {{ margin: 5px 0; }}
            </style></head><body>
            <div class="header">
            <h1>Statistiche Assenze Docenti</h1>
            <p>Data report: {datetime.now().strftime('%d/%m/%Y')}</p>
            </div>
            <h3>Riepilogo Percentuale Assenze</h3>
            <div class="stat-box"><strong>Settimanale:</strong> {w}%</div>
            <div class="stat-box"><strong>Mensile:</strong> {m}%</div>
            <div class="stat-box"><strong>Totale:</strong> {t}%</div>
            <br>
            <p><strong>Totale Monte Ore Docenti (Sommatoria):</strong> {d}</p>

            <h3>Dettaglio Ore Assenza (Totale Periodo)</h3>
            <table>
            <tr><th>Docente</th><th>Ore Totali Assenza</th></tr>
            {rows_html}
            </table>

            <div style='text-align:center; margin-top:20px'>© Mario Carta</div>
            <script>window.print()</script></body></html>"""
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix='.html', mode='w', encoding='utf-8') as f:
                f.write(html)
            self.temp_files.append(f.name)
            webbrowser.open('file://' + f.name)
        except Exception as e:
            logger.error(f"Errore stampa statistiche: {e}")
            messagebox.showerror("Errore", str(e))

    def export_db_file(self):
        """Esporta il database su file."""
        path = filedialog.asksaveasfilename(defaultextension=".sqlite",
                                             filetypes=[("SQLite DB", "*.sqlite")])
        if path:
            try:
                self.db.export_db(path)
                messagebox.showinfo("Export", "Database esportato correttamente!")
            except Exception as e:
                messagebox.showerror("Export", str(e))

    def import_db_file(self):
        """Importa un database da file esterno."""
        path = filedialog.askopenfilename(filetypes=[("SQLite DB", "*.sqlite")])
        if path:
            confirm = messagebox.askyesno("Conferma",
                                           "Se importi un database, i dati attuali verranno sovrascritti. Continuare?")
            if confirm:
                try:
                    self.db.import_db(path)
                    if self.db.load_state(self.dm):
                        self.refresh_labels()
                        self.txt_absent.delete("1.0", "end")
                        self.txt_short_leave.delete("1.0", "end")
                        self.txt_classes_absent.delete("1.0", "end")
                        self.entry_total_teachers.delete(0, "end")
                        self.entry_total_teachers.insert(0, str(self.dm.total_teachers))
                        self.update_stats_display()
                        messagebox.showinfo("Import", "Database importato con successo!")
                    else:
                        messagebox.showerror("Errore", "Impossibile caricare il database importato.")
                except Exception as e:
                    messagebox.showerror("Errore Import", str(e))

    def clear_table(self):
        """Svuota la tabella dei risultati."""
        for item in self.tree.get_children():
            self.tree.delete(item)

    def run_calculation(self):
        """Esegue il calcolo delle sostituzioni."""
        self.clear_table()
        day = self.combo_day.get()
        raw_input = self.entry_teacher.get().strip()

        # Teachers explicitly requested for substitution (UI Display)
        ui_targets_set = set()
        if raw_input:
            ui_targets_set = {to_slug(x) for x in re.split(r'[,;]+', raw_input) if len(x) > 2}

        raw_bl = self.txt_absent.get("1.0", "end")
        # Teachers explicitly listed as absent (Statistics)
        stats_targets_list = [to_slug(x) for x in re.split(r'[\n,]+', raw_bl) if len(x) > 2]
        stats_targets_set = set(stats_targets_list)

        # Union of both for processing
        processing_targets = list(ui_targets_set.union(stats_targets_set))

        if not processing_targets:
            return messagebox.showerror("Errore", "Inserisci almeno un docente (da sostituire o assente)!")

        # Blacklist for engine is the list of ACTUALLY absent teachers
        blacklist = list(stats_targets_list)

        short_leaves_map = {}
        raw_short = self.txt_short_leave.get("1.0", "end")
        for line in raw_short.splitlines():
            for entry in line.strip().split(','):
                parts = entry.strip().split()
                if len(parts) >= 2:
                    s_slug = to_slug(parts[0])
                    hours = [int(p) for p in parts[1:] if p.isdigit()]
                    if s_slug and hours:
                        if s_slug not in short_leaves_map:
                            short_leaves_map[s_slug] = []
                        short_leaves_map[s_slug].extend(hours)

        # Teachers to be replaced treated as 'busy'
        for t in processing_targets:
            if t not in short_leaves_map and t not in blacklist:
                blacklist.append(t)

        raw_cls_abs = self.txt_classes_absent.get("1.0", "end")
        absent_classes_slugs = [normalize_class(x) for x in re.split(r'[\n,]+', raw_cls_abs) if len(x) > 1]

        results = self.engine.find_substitutions(day, processing_targets, blacklist,
                                                  short_leaves_map, absent_classes_slugs)
        # Sort results by Class and Hour
        results.sort(key=lambda x: (x[0], x[1]))

        # STATS ACCUMULATION
        try:
            tot = int(self.entry_total_teachers.get())
            self.dm.total_teachers = tot

            stats_valid_set = stats_targets_set.union(short_leaves_map.keys())

            new_details = {}
            if results:
                for r in results:
                    if len(r) >= 6:
                        t_slugs = str(r[5]).split(',')
                        for ts in t_slugs:
                            if ts and ts in stats_valid_set:
                                new_details[ts] = new_details.get(ts, 0) + 1

            if new_details:
                self.db.accumulate_daily_stats(new_details, tot)
                self.update_stats_display()
        except Exception as e:
            logger.error(f"Errore aggiornamento statistiche: {e}")

        if not results:
            messagebox.showinfo("Info", "Nessun risultato trovato.")

        # UI FILTER: se nessun filtro UI specificato, mostra tutti i risultati
        show_all = len(ui_targets_set) == 0
        for r in results:
            if len(r) >= 6:
                t_slugs = str(r[5]).split(',')
                if show_all or any(ts in ui_targets_set for ts in t_slugs):
                    self.tree.insert("", "end", values=r[:4], tags=(r[4],))

    def export_excel(self):
        """Esporta la tabella risultati in formato Excel."""
        if not HAS_PANDAS:
            return messagebox.showerror("Errore", "Libreria pandas non trovata.")
        data = []
        for item in self.tree.get_children():
            data.append(self.tree.item(item)['values'])
        if not data:
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return

        # Lettura dati dal thread principale (thread-safety)
        abs_text = self.txt_absent.get("1.0", "end").strip()
        short_text = self.txt_short_leave.get("1.0", "end").strip()
        cls_text = self.txt_classes_absent.get("1.0", "end").strip()
        raw_input = self.entry_teacher.get().strip()

        # Calc Off Site (escludendo sia i docenti assenti per intera giornata sia quelli con permesso breve)
        ui_targets_set = {to_slug(x) for x in re.split(r'[,;]+', raw_input) if len(x) > 2} if raw_input else set()
        stats_targets_set = {to_slug(x) for x in re.split(r'[\n,]+', abs_text) if len(x) > 2}
        short_leaves_set = set()
        for line in short_text.splitlines():
            for entry in line.strip().split(','):
                parts = entry.strip().split()
                if len(parts) >= 2 and any(p.isdigit() for p in parts[1:]):
                    s_slug = to_slug(parts[0])
                    if s_slug:
                        short_leaves_set.add(s_slug)
        off_site_set = ui_targets_set - stats_targets_set - short_leaves_set
        off_site_txt = ", ".join(sorted(off_site_set))

        def worker():
            try:
                df_main = pd.DataFrame(data, columns=["Classe", "Ora", "Situazione", "Sostituzione"])
                summary_data = [
                    ["Docenti Assenti", abs_text],
                    ["Docenti Fuori Sede", off_site_txt],
                    ["Permessi Brevi", short_text],
                    ["Classi Assenti", cls_text],
                    ["", ""]
                ]
                df_summary = pd.DataFrame(summary_data, columns=["Tipo", "Valore"])
                with pd.ExcelWriter(path, engine='openpyxl') as writer:
                    df_summary.to_excel(writer, index=False, header=False, startrow=0)
                    df_main.to_excel(writer, index=False, startrow=len(summary_data) + 2)
                self.after(0, lambda: messagebox.showinfo("Fatto", "Excel salvato."))
                logger.info(f"Excel esportato: {path}")
            except Exception as e:
                logger.error(f"Errore esportazione Excel: {e}", exc_info=True)
                self.after(0, lambda: messagebox.showerror("Errore", str(e)))

        threading.Thread(target=worker, daemon=True).start()

    def ask_export_word_mode(self):
        """Dialog modale per scegliere se esportare Word a colori o in B/N."""
        choice = [None]
        dialog = tk.Toplevel(self)
        dialog.title("Esportazione Word")
        dialog.geometry("460x230")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        dialog.configure(bg="#F8F9FA")

        # Centra la finestra
        dialog.update_idletasks()
        try:
            pw, ph = self.winfo_width(), self.winfo_height()
            px, py = self.winfo_rootx(), self.winfo_rooty()
            dx = px + (pw - 460) // 2
            dy = py + (ph - 230) // 2
            dialog.geometry(f"+{max(0, dx)}+{max(0, dy)}")
        except Exception:
            pass

        tk.Label(dialog, text="Esportazione Documento Word",
                 font=("Helvetica", 14, "bold"), bg="#F8F9FA", fg="#2c3e50").pack(pady=(18, 6))
        tk.Label(dialog, text="Come desideri formattare le sostituzioni nel file Word?",
                 font=("Segoe UI", 11), bg="#F8F9FA", fg="#555555").pack(pady=(0, 16))

        btn_frame = tk.Frame(dialog, bg="#F8F9FA")
        btn_frame.pack(pady=5)

        def choose(val):
            choice[0] = val
            dialog.destroy()

        btn_color = tk.Button(btn_frame, text="🎨  A Colori\n(come Anteprima PDF)",
                              font=("Segoe UI", 10, "bold"),
                              bg="#E3F2FD", fg="#1565C0", activebackground="#BBDEFB",
                              relief="groove", padx=16, pady=8, cursor="hand2",
                              command=lambda: choose("color"))
        btn_color.pack(side="left", padx=10)

        btn_bw = tk.Button(btn_frame, text="📄  Bianco e Nero\n(Stampa B/N standard)",
                           font=("Segoe UI", 10, "bold"),
                           bg="#ECEFF1", fg="#37474F", activebackground="#CFD8DC",
                           relief="groove", padx=16, pady=8, cursor="hand2",
                           command=lambda: choose("bw"))
        btn_bw.pack(side="left", padx=10)

        btn_cancel = tk.Button(dialog, text="Annulla", font=("Segoe UI", 10),
                               bg="#F5F5F5", fg="#7f8c8d", relief="flat",
                               command=lambda: choose(None))
        btn_cancel.pack(pady=(14, 10))

        dialog.bind("<Escape>", lambda e: choose(None))
        self.wait_window(dialog)
        return choice[0]

    def export_word(self):
        """Esporta la tabella risultati in formato Word (.docx)."""
        if not HAS_DOCX:
            return messagebox.showerror("Errore", "python-docx non trovata.")
        if not self.tree.get_children():
            return messagebox.showinfo("Info",
                                        "Nessuna riga di sostituzione presente nella tabella da esportare.")

        # Selezione modalità (A Colori o Bianco e Nero)
        mode = self.ask_export_word_mode()
        if not mode:
            return

        path = filedialog.asksaveasfilename(defaultextension=".docx",
                                             filetypes=[("Word files", "*.docx")])
        if not path:
            return

        # Lettura dati dal thread principale (thread-safety)
        teacher_txt = self.entry_teacher.get().upper()
        day_txt = self.combo_day.get()

        # Raccoglie valori e tag per ciascuna riga della tabella
        tree_items_data = []
        for item in self.tree.get_children():
            vals = self.tree.item(item)['values']
            tags = self.tree.item(item)['tags']
            tag = tags[0] if tags else ""
            tree_items_data.append((vals, tag))

        abs_txt = self.txt_absent.get("1.0", "end").strip()
        short_txt = self.txt_short_leave.get("1.0", "end").strip()
        cls_abs_txt = self.txt_classes_absent.get("1.0", "end").strip()
        logo_path = self.logo_path

        # Calc Off Site (escludendo sia i docenti assenti per intera giornata sia quelli con permesso breve)
        raw_input = self.entry_teacher.get().strip()
        ui_targets_set = {to_slug(x) for x in re.split(r'[,;]+', raw_input) if len(x) > 2} if raw_input else set()
        stats_targets_set = {to_slug(x) for x in re.split(r'[\n,]+', abs_txt) if len(x) > 2}
        short_leaves_set = set()
        for line in short_txt.splitlines():
            for entry in line.strip().split(','):
                parts = entry.strip().split()
                if len(parts) >= 2 and any(p.isdigit() for p in parts[1:]):
                    s_slug = to_slug(parts[0])
                    if s_slug:
                        short_leaves_set.add(s_slug)
        off_site_set = ui_targets_set - stats_targets_set - short_leaves_set
        off_site_txt = ", ".join(sorted(off_site_set))

        def worker():
            try:
                doc = Document()
                # Access header safely
                header = doc.sections[0].header
                if not header.paragraphs:
                    p = header.add_paragraph()
                else:
                    p = header.paragraphs[0]

                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.text = "I.P.S.A.R - I.P.S.E.O.A Alberghiero Sassari"

                if logo_path and os.path.exists(logo_path):
                    p_logo = header.add_paragraph()
                    p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_logo.add_run().add_picture(logo_path, width=Inches(1.2))

                doc.add_heading(f"Docenti Assenti - {datetime.now().strftime('%d/%m/%Y')}", 0)
                doc.add_paragraph(f"Docenti: {teacher_txt} | Giorno: {day_txt}")

                # 1. Tabella Riepilogo Assenze
                table_s = doc.add_table(rows=2, cols=4)
                table_s.style = 'Table Grid'
                hdr_s = table_s.rows[0].cells
                hdr_s_titles = ["Docenti Assenti", "Docenti Fuori Sede", "Permessi Brevi", "Classi Assenti"]
                for i, title in enumerate(hdr_s_titles):
                    hdr_s[i].text = title
                    p = hdr_s[i].paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in p.runs:
                        run.bold = True
                        run.font.size = Pt(10)
                    if mode == "color":
                        set_docx_cell_shading(hdr_s[i], "#E3F2FD")
                    else:
                        set_docx_cell_shading(hdr_s[i], "#EEEEEE")

                row_s = table_s.rows[1].cells
                row_s_vals = [abs_txt, off_site_txt, short_txt, cls_abs_txt]
                for i, val in enumerate(row_s_vals):
                    row_s[i].text = val
                    for p in row_s[i].paragraphs:
                        for run in p.runs:
                            run.font.size = Pt(9.5)

                for row in table_s.rows:
                    for cell in row.cells:
                        cell.width = Inches(1.7)

                doc.add_paragraph("\n")

                # 2. Tabella Principale Sostituzioni
                table = doc.add_table(rows=1, cols=4)
                table.style = 'Table Grid'
                hdr = table.rows[0].cells
                hdr_titles = ["Classe", "Ora", "Situazione", "Sostituzione"]
                for i, title in enumerate(hdr_titles):
                    hdr[i].text = title
                    p = hdr[i].paragraphs[0]
                    if i in (0, 1):
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in p.runs:
                        run.bold = True
                        run.font.size = Pt(10.5)
                        if mode == "color":
                            run.font.color.rgb = RGBColor(255, 255, 255)
                    if mode == "color":
                        set_docx_cell_shading(hdr[i], "#1565C0")
                    else:
                        set_docx_cell_shading(hdr[i], "#E0E0E0")

                for vals, tag in tree_items_data:
                    row = table.add_row().cells
                    bg_color = None
                    if mode == "color":
                        if tag == "coperta":
                            bg_color = "#A5D6A7"
                        elif tag in Configuration.COLORS:
                            bg_color = Configuration.COLORS[tag]

                    for i in range(4):
                        val_str = str(vals[i]) if i < len(vals) else ""
                        row[i].text = val_str
                        p = row[i].paragraphs[0]
                        if i == 0:
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            for run in p.runs:
                                run.bold = True
                                run.font.size = Pt(10)
                        elif i == 1:
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            for run in p.runs:
                                run.font.size = Pt(10)
                        else:
                            for p_sub in row[i].paragraphs:
                                for run in p_sub.runs:
                                    run.font.size = Pt(9.5)

                        if bg_color:
                            set_docx_cell_shading(row[i], bg_color)

                # Larghezze ottimali per colonna
                col_widths = [Inches(1.0), Inches(0.6), Inches(2.2), Inches(3.0)]
                for row in table.rows:
                    for i, w in enumerate(col_widths):
                        row.cells[i].width = w

                doc.add_paragraph("\n© Mario Carta", style='Quote')
                doc.save(path)
                mode_label = "a colori" if mode == "color" else "in bianco e nero"
                self.after(0, lambda: messagebox.showinfo("Fatto",
                    f"Documento Word ({mode_label}) salvato con successo."))
                logger.info(f"Word esportato ({mode}): {path}")
            except Exception as e:
                err_details = traceback.format_exc()
                logger.error(f"Errore esportazione Word: {err_details}")
                self.after(0, lambda: messagebox.showerror("Errore Esportazione Word",
                    f"Si è verificato un errore:\n\n{str(e)}\n\nDettagli:\n{err_details[:300]}..."))

        threading.Thread(target=worker, daemon=True).start()

    def print_preview(self):
        """Genera e apre nel browser l'anteprima di stampa PDF."""
        txt_ass = self.txt_absent.get("1.0", "end").strip().replace('\n', '<br>')
        txt_per = self.txt_short_leave.get("1.0", "end").strip().replace('\n', '<br>')
        txt_cla = self.txt_classes_absent.get("1.0", "end").strip().replace('\n', '<br>')

        # Calc Off Site (escludendo sia i docenti assenti per intera giornata sia quelli con permesso breve)
        raw_input = self.entry_teacher.get().strip()
        ui_targets_set = {to_slug(x) for x in re.split(r'[,;]+', raw_input) if len(x) > 2} if raw_input else set()
        raw_bl = self.txt_absent.get("1.0", "end")
        stats_targets_set = {to_slug(x) for x in re.split(r'[\n,]+', raw_bl) if len(x) > 2}
        raw_short = self.txt_short_leave.get("1.0", "end")
        short_leaves_set = set()
        for line in raw_short.splitlines():
            for entry in line.strip().split(','):
                parts = entry.strip().split()
                if len(parts) >= 2 and any(p.isdigit() for p in parts[1:]):
                    s_slug = to_slug(parts[0])
                    if s_slug:
                        short_leaves_set.add(s_slug)
        off_site_set = ui_targets_set - stats_targets_set - short_leaves_set
        off_site_txt = ", ".join(sorted(off_site_set)).replace('\n', '<br>')

        html = f"""<html><head><style>
        body {{ font-family: sans-serif; padding: 20px; }}
        .header {{ text-align: center; border-bottom: 3px solid #1565C0; padding-bottom: 10px; margin-bottom: 20px; }}
        .logo {{ max-height: 100px; display: block; margin: 0 auto; }}
        h1 {{ color: #1565C0; margin: 5px 0; }}
        table {{ width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 14px; width:100%; }}
        th, td {{ border: 1px solid #999; padding: 8px; text-align: left; vertical-align: top; }}
        th {{ background-color: #EEE; }}
        .summary-table th {{ background-color: #e3f2fd; text-align: center; }}
        """
        for k, v in Configuration.COLORS.items():
            html += f"tr.{k} td {{ background-color: {v}; }}\n"

        html += f"""
        </style></head><body>
        <div class="header">
            <img src="{os.path.abspath(self.logo_path) if self.logo_path and os.path.exists(self.logo_path) else ''}" class="logo">
            <h1>I.P.S.A.R. - I.P.S.E.O.A. SASSARI</h1>
            <h3>Docenti Assenti: {self.entry_teacher.get().upper()} - {self.combo_day.get()}</h3>
            <small>{datetime.now().strftime('%d/%m/%Y')}</small>
        </div>

        <table class="summary-table">
            <tr><th width="25%">Docenti Assenti</th><th width="25%">Docenti Fuori Sede</th><th width="25%">Permessi Brevi</th><th width="25%">Classi Assenti</th></tr>
            <tr><td>{txt_ass}</td><td>{off_site_txt}</td><td>{txt_per}</td><td>{txt_cla}</td></tr>
        </table>
        <br>
        <table><tr><th>Classe</th><th>Ora</th><th>Situazione</th><th>Sostituzione</th></tr>"""

        for item in self.tree.get_children():
            v = self.tree.item(item)['values']
            t = self.tree.item(item)['tags']
            tag = t[0] if t else ""
            sub = str(v[3]).replace("\n", "<br>")
            html += f"<tr class='{tag}'><td><b>{v[0]}</b></td><td style='text-align:center'>{v[1]}</td><td>{v[2]}</td><td>{sub}</td></tr>"

        html += "</table><div style='text-align:center; margin-top:20px'>© Mario Carta</div><script>window.print()</script></body></html>"

        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix='.html', mode='w', encoding='utf-8') as f:
                f.write(html)
            self.temp_files.append(f.name)
            webbrowser.open('file://' + f.name)
        except Exception as e:
            logger.error(f"Errore anteprima stampa: {e}")
            messagebox.showerror("Errore", str(e))
