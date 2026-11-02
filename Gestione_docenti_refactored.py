import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import re
import os
import sys
import traceback
import webbrowser
import tempfile
import csv
import io
import threading
import sqlite3
import json
import shutil
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Tuple, Any

# --- LIBRARIES ---
try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    import lxml
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

def set_docx_cell_shading(cell, color_hex):
    """Applica un colore di sfondo a una cella di tabella Word (.docx)."""
    if not color_hex: return
    fill = color_hex.lstrip('#').upper()
    try:
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill}"/>')
        cell._tc.get_or_add_tcPr().append(shading)
    except Exception:
        pass

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

# --- CONFIGURATION ---

class Configuration:
    SUBJECT_BLACKLIST = [
        "ITALIANO", "STORIA", "INGLESE", "MATEMATICA", "RELIGIONE", "ED.FISICA",
        "SOSTEGNO", "DIRITTO", "ECONOMIA", "INFORMATICA", "SCIENZE", "CHIMICA",
        "FISICA", "GEO", "LETTERE", "LABOR.", "LABORATORIO", "ENOGASTRONOMIA",
        "CUCINA", "SALA", "VENDITA", "ACCOGLIENZA", "TURISTICA", "ALIMENTAZIONE",
        "FRANCESE", "TEDESCO", "SPAGNOLO", "TIC", "TTP", "CONV"
    ]

    DAYS = [None, "LUNEDÌ", "MARTEDÌ", "MERCOLEDÌ", "GIOVEDÌ", "VENERDÌ", "SABATO"]

    COLORS = {
        'compresenza': '#B2DFDB', 
        'disp_class': '#FFF9C4', 
        'potenziamento': '#FFE0B2',
        'pagamento': '#C8E6C9', 
        'scoperta': '#FFCDD2', 
        'alternativa': '#E1BEE7'
    }
    
    THEME = {
        'primary': '#2c3e50', 
        'secondary': '#ecf0f1', 
        'accent': '#2980b9', 
        'text': '#2c3e50',
        'bg': '#f4f6f7',
        'btn_bg': '#F0F0F0', 
        'btn_fg': '#000000'
    }
    
    DB_NAME = "sostituzioni_db.sqlite"

    @staticmethod
    def get_db_path():
        # [FIX 2.8] Percorso scrivibile per ogni piattaforma
        if sys.platform == 'darwin':
            # Use Application Support on Mac to ensure writability
            lib_path = os.path.expanduser('~/Library/Application Support/GestioneSostituzioni')
        elif sys.platform == 'win32':
            # Use APPDATA on Windows to ensure writability (anche da PyInstaller)
            lib_path = os.path.join(os.environ.get('APPDATA', os.path.expanduser('~')), 'GestioneSostituzioni')
        else:
            # Linux/other
            lib_path = os.path.expanduser('~/.config/GestioneSostituzioni')
        if not os.path.exists(lib_path):
            os.makedirs(lib_path, exist_ok=True)
        return os.path.join(lib_path, Configuration.DB_NAME)

    @staticmethod
    def resource_path(relative_path: str) -> str:
        try:
            base_path = sys._MEIPASS
        except AttributeError:
            base_path = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_path, relative_path)

# --- UTILS ---

def clean_str(s: str) -> str:
    return s.strip().replace('"', '') if s else ""

def to_slug(str_val: str) -> str:
    # [FIX 2.5] Include prime lettere del nome per disambiguare omonimie
    if not str_val: return ""
    no_parens = re.sub(r'\([^)]*\)', ' ', str_val)
    clean = re.sub(r'[^A-Z0-9\s]', ' ', no_parens.upper())
    parts = clean.split()
    if not parts: return ""
    surname = parts[0]
    prefixes = ["DE", "DI", "LO", "LA", "IL", "DEL", "DELLA", "DEGLI", "DAL", "SAN", "SANTA", "DA"]
    if len(parts) > 1 and surname in prefixes:
        surname += parts[1]
        # Aggiungi iniziale del nome se disponibile (dopo il cognome composto)
        if len(parts) > 2:
            surname += parts[2][:3]
    elif len(parts) > 1:
        # Aggiungi prime 3 lettere del nome per disambiguare omonimie
        surname += parts[1][:3]
    return surname

def is_subject(word: str) -> bool:
    # [FIX 2.4] Match esatto per evitare falsi positivi su cognomi reali (es. Francese, Converso)
    w = word.upper().replace('.', '').replace(',', '')
    EXACT_MATCH = {"ITALIANO", "STORIA", "INGLESE", "MATEMATICA", "RELIGIONE",
                   "SOSTEGNO", "DIRITTO", "ECONOMIA", "INFORMATICA", "SCIENZE",
                   "CHIMICA", "FISICA", "LETTERE", "LABORATORIO", "ENOGASTRONOMIA",
                   "CUCINA", "SALA", "VENDITA", "ACCOGLIENZA", "TURISTICA",
                   "ALIMENTAZIONE", "FRANCESE", "TEDESCO", "SPAGNOLO", "TIC", "TTP"}
    PREFIX_MATCH = {"EDFISICA", "LABOR", "CONV", "GEO"}
    return w in EXACT_MATCH or any(w.startswith(p) for p in PREFIX_MATCH)

def normalize_class(name: str) -> str:
    # [FIX 2.3] Solo le sigle esplicite di biennio vengono espanse,
    # evitando di trasformare classi come 1EN, 2SV, 3AT in stringhe errate
    if not name: return ""
    n = name.upper().strip()
    is_vendita = ("VENDITA" in n) or ("SALA" in n)
    clean_base = re.sub(r'[\(\)\s]*(LAB|SALA|VENDITA)[\(\)\s]*', '', n).strip()
    clean_base = re.sub(r'\s+', '', clean_base)
    BIENNIO_SUFFIXES = {"AB", "BB", "CB", "DB", "EB", "FB"}
    m = re.match(r'^(\d)([A-Z]{2})$', clean_base)
    if m and m.group(2) in BIENNIO_SUFFIXES:
        clean_base = m.group(1) + m.group(2)[0] + "BIENNIO"
    if "BIENNIO" in clean_base:
        clean_base = clean_base.replace("BIENNIO", "ABIENNIO").replace("AA", "A")
    is_biennio = bool(re.match(r'^[12]', clean_base)) or bool(re.match(r'^[12]', n))
    if is_vendita and is_biennio:
        return f"{clean_base} VENDITA"
    return clean_base

def get_hour(text: str) -> int:
    # [FIX 2.2] Supporta sia ore scolastiche ordinali (1-8) che fasce orarie reali (08:00-15:00)
    if not text: return 0
    s = text.strip()
    # Tentativo 1: Numero semplice 1-8 (ora scolastica diretta)
    if s.isdigit():
        h = int(s)
        if 1 <= h <= 8: return h
    # Tentativo 2: Formato orario HH:MM o HH.MM
    match = re.search(r'(^|\s|:)0?(\d{1,2})[:\.]', s)
    if not match: match = re.search(r'^0?(\d{1,2})$', s)
    if match:
        try:
            val = match.group(2) if len(match.groups()) > 1 and match.group(2) else match.group(1)
            h = int(val)
            if 8 <= h <= 15: return h - 7
        except: return 0
    return 0

# --- DATA MANAGER ---

class DataManager:
    def __init__(self):
        self.db_main = {}
        self.db_pers = {}
        self.db_sos = {}   
        self.db_alt = {d: {} for d in Configuration.DAYS if d}
        self.db_pot = {}
        self.db_pay = {}
        self.text_absent = ""
        self.text_short = ""
        self.text_classes = ""
        self.preview_data = {} 
        
        # Stats
        self.total_teachers = 0

    def _get_content(self, source: str, is_path: bool) -> Tuple[List[str], str]:
        # [FIX 2.5b] Rilevamento automatico encoding con fallback a cascata
        if is_path:
            encodings = ['utf-8-sig', 'utf-8', 'cp1252', 'ISO-8859-1']
            for enc in encodings:
                try:
                    with open(source, 'r', encoding=enc) as f:
                        content = f.read()
                        return content.splitlines(), content
                except (UnicodeDecodeError, UnicodeError):
                    continue
                except Exception:
                    return [], ""
            return [], ""
        else:
            return source.splitlines(), source

    def read_csv_rows(self, lines: List[str]) -> List[List[str]]:
        if not lines: return []
        try:
            # Try to detect dialect
            sample = "\n".join(lines[:5])
            try: dialect = csv.Sniffer().sniff(sample, delimiters=';,')
            except: dialect = None
            
            if dialect: reader = csv.reader(lines, dialect)
            else:
                # Fallback default
                sep = ';' if lines[0].count(';') >= lines[0].count(',') else ','
                reader = csv.reader(lines, delimiter=sep)
            
            return list(reader)
        except Exception as e:
            print(f"CSV Read Error: {e}")
            # Ultra-fallback
            sep = ';' if lines[0].count(';') >= lines[0].count(',') else ','
            return [line.split(sep) for line in lines]

    def parse_main(self, source: str, is_path: bool = True) -> int:
        self.db_main = {}
        lines, raw = self._get_content(source, is_path)
        self.preview_data['main'] = raw 
        if not lines: return 0
        
        rows = self.read_csv_rows(lines)
        
        curr_class = None; last_h = 0; day_map = {}; count = 0
        for row in rows:
            # Reconstruct checks on "line" using joined row for regex or strings, 
            # BUT we already have cols. 
            # Original code checked `line.upper()`. We can join row back for that or check cols.
            # Ideally checking cols is safer.
            
            # Helper to match original logic which inspected the full line string
            line_str = " ".join(row) 
            upper = line_str.upper(); cols = [clean_str(c) for c in row]
            
            if "LUNED" in upper or "MARTED" in upper:
                for i, c in enumerate(cols):
                    for d in Configuration.DAYS:
                        if d and d[:3] in c.upper(): day_map[i] = d; break
            if "CLASSE" in upper:
                is_vendita_or_sala = ("VENDITA" in upper) or ("SALA" in upper)
                m = re.search(r'CLASSE\s+([^,;\(]+)', line_str, re.IGNORECASE)
                if m:
                    base_class = normalize_class(m.group(1))
                    is_biennio = len(base_class) >= 2 and base_class[0] in "12"
                    if is_biennio and (is_vendita_or_sala or base_class in self.db_main):
                        curr_class = f"{base_class} VENDITA" if not base_class.endswith(" VENDITA") else base_class
                    else:
                        curr_class = base_class
                    if curr_class not in self.db_main:
                        self.db_main[curr_class] = {}
                        count += 1
                    last_h = 0
            
            if curr_class and len(cols) > 1:
                # Check column 0 for hour
                h = get_hour(cols[0]); use_def = len(day_map) == 0
                if h > 0:
                    last_h = h
                    for i in range(1, len(cols)):
                        if i >= len(cols): break
                        day = Configuration.DAYS[i] if use_def else day_map.get(i)
                        if day and len(cols[i]) > 1:
                            if day not in self.db_main[curr_class]: self.db_main[curr_class][day] = {}
                            self.db_main[curr_class][day][h] = {'text': cols[i], 'blocked': []}
                elif "COMPR" in upper and last_h > 0:
                    for i in range(1, len(cols)):
                        if i >= len(cols): break
                        day = Configuration.DAYS[i] if use_def else day_map.get(i)
                        if day and len(cols[i]) > 1:
                            if day in self.db_main[curr_class] and last_h in self.db_main[curr_class][day]:
                                self.db_main[curr_class][day][last_h]['text'] += f" (Compr: {cols[i]})"
                                # [FIX 2.1] Compresenti separati dai 'blocked': i compresenti
                                # sono docenti disponibili come sostituti interni, non da bloccare
                                toks = [t for t in re.split(r'[\s\-\/]+', cols[i].upper()) if len(t) > 1 and not is_subject(t)]
                                if 'compresenti' not in self.db_main[curr_class][day][last_h]:
                                    self.db_main[curr_class][day][last_h]['compresenti'] = []
                                self.db_main[curr_class][day][last_h]['compresenti'].extend(toks)
        return count

    def parse_pers(self, source: str, key='pers', is_path: bool = True) -> int:
        target = {}
        lines, raw = self._get_content(source, is_path)
        self.preview_data[key] = raw
        if not lines: return 0
        
        rows = self.read_csv_rows(lines); loaded = 0
        ignore = ["DOCENTE", "MATERIA", "ORA"] + Configuration.SUBJECT_BLACKLIST
        for row in rows:
            cols = [clean_str(c) for c in row]
            if not cols: continue
            c0 = cols[0].strip()
            if len(c0) > 2 and not any(c0.upper().startswith(x) for x in ignore):
                curr = to_slug(c0)
                if curr not in target: target[curr] = {}
                for d in range(6):
                    day = Configuration.DAYS[d+1]; t_idx = 1 + (d*2); c_idx = t_idx + 1
                    if c_idx < len(cols):
                        h = get_hour(cols[t_idx])
                        if h > 0 and len(cols[c_idx]) > 1:
                            if day not in target[curr]: target[curr][day] = {}
                            target[curr][day][h] = normalize_class(cols[c_idx]); loaded += 1
        
        if key == 'pers': self.db_pers = target
        elif key == 'sos': self.db_sos = target
        return loaded

    def parse_alt(self, source: str, is_path: bool = True) -> int:
        self.db_alt = {d: {} for d in Configuration.DAYS if d}
        lines, raw = self._get_content(source, is_path)
        self.preview_data['alt'] = raw
        if not lines: return 0
        
        rows = self.read_csv_rows(lines); loaded = 0
        for row in rows:
            cols = [clean_str(c) for c in row]
            for d in range(6):
                day = Configuration.DAYS[d+1]; t_idx = d*2; c_idx = t_idx + 1
                if c_idx < len(cols):
                    h = get_hour(cols[t_idx]); val = cols[c_idx]
                    if h > 0 and len(val) > 2:
                        parts = val.strip().split()
                        if len(parts) >= 2:
                            cls = normalize_class(parts[0])
                            teacher = " ".join([p for p in parts[1:] if not is_subject(p)])
                            if h not in self.db_alt[day]: self.db_alt[day][h] = {}
                            self.db_alt[day][h][cls] = teacher; loaded += 1
        return loaded

    def parse_generic(self, source: str, key='pot', is_path: bool = True) -> int:
        target_db = {d: {h: [] for h in range(1,9)} for d in Configuration.DAYS if d}
        lines, raw = self._get_content(source, is_path)
        self.preview_data[key] = raw
        if not lines: return 0
        
        rows = self.read_csv_rows(lines); loaded = 0
        for row in rows:
            cols = [clean_str(c) for c in row]
            for d in range(6):
                day = Configuration.DAYS[d+1]; t_idx = d*2; n_idx = t_idx + 1
                if n_idx < len(cols):
                    h = get_hour(cols[t_idx]); raw = cols[n_idx]
                    if h > 0 and len(raw) > 2:
                        parts = re.split(r'\s+', raw)
                        for n in parts:
                            if len(n) > 2 and not is_subject(n):
                                target_db[day][h].append(n); loaded += 1
        if key == 'pot': self.db_pot = target_db
        else: self.db_pay = target_db
        return loaded

# --- DATABASE HANDLER ---

class DatabaseHandler:
    def __init__(self):
        self.db_path = Configuration.get_db_path()
        
        # [MIGRATION] If old local DB exists and new one doesn't, copy it
        bundled_db = Configuration.resource_path(Configuration.DB_NAME)
        if os.path.exists(bundled_db) and not os.path.exists(self.db_path):
            try: shutil.copy2(bundled_db, self.db_path)
            except: pass
        
    def get_connection(self):
        return sqlite3.connect(self.db_path)

    def setup(self):
        with self.get_connection() as conn:
            c = conn.cursor()
            c.execute('''CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT)''')
            c.execute('''CREATE TABLE IF NOT EXISTS data_store (key TEXT PRIMARY KEY, json_data TEXT)''')
            c.execute('''CREATE TABLE IF NOT EXISTS preview_store (key TEXT PRIMARY KEY, content TEXT)''')
            # New Stats table
            c.execute('''CREATE TABLE IF NOT EXISTS daily_stats (date TEXT PRIMARY KEY, absent_count INTEGER, total_teachers INTEGER, details TEXT)''')
            # Attempt migration for existing DBs
            try: c.execute("ALTER TABLE daily_stats ADD COLUMN details TEXT")
            except: pass
            conn.commit()

    def accumulate_daily_stats(self, new_details_dict, tot_teachers):
        # [FIX 1.4] Sostituire i dati del giorno corrente anziché accumularli,
        # evitando la moltiplicazione delle ore ad ogni pressione di PROCEDI
        today = datetime.now().strftime("%Y-%m-%d")
        
        count = len(new_details_dict)
        details_json = json.dumps(new_details_dict)
        
        with self.get_connection() as conn:
            c = conn.cursor()
            c.execute("INSERT OR REPLACE INTO daily_stats (date, absent_count, total_teachers, details) VALUES (?, ?, ?, ?)", 
                      (today, count, tot_teachers, details_json))
            conn.commit()
            
    def get_stats(self):
        # Returns (week_avg, month_avg, total_avg, total_hours, teacher_details_map)
        with self.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT date, absent_count, total_teachers, details FROM daily_stats")
            rows = c.fetchall()
        
        if not rows: return (0,0,0,0, {})

        now = datetime.now()
        start_week = (now - timedelta(days=now.weekday())).strftime("%Y-%m-%d")
        start_month = now.strftime("%Y-%m-01")
        
        w_n, w_d = 0, 0
        m_n, m_d = 0, 0
        t_n, t_d = 0, 0
        
        agg_details = {} # "Teacher": hours
        
        for r in rows:
            dt, count, total, det_str = r
            if total > 0:
                t_n += count; t_d += total
                if dt >= start_week: w_n += count; w_d += total
                if dt >= start_month: m_n += count; m_d += total
            
            # Aggregate hours
            if det_str:
                try:
                    d_map = json.loads(det_str)
                    for t_name, h_count in d_map.items():
                        agg_details[t_name] = agg_details.get(t_name, 0) + h_count
                except: pass
        
        def safe_pct(n, d): return round((n/d)*100, 2) if d > 0 else 0
        return (safe_pct(w_n, w_d), safe_pct(m_n, m_d), safe_pct(t_n, t_d), t_d, agg_details)

    def reset_stats(self):
        with self.get_connection() as conn:
            conn.execute("DELETE FROM daily_stats")
            conn.commit()

    def save_state(self, dm: DataManager):
        with self.get_connection() as conn:
            c = conn.cursor()
            c.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", ('absent', ''))
            c.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", ('short', ''))
            c.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", ('classes', ''))
            c.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", ('total_teachers', str(dm.total_teachers)))
            
            c.execute("INSERT OR REPLACE INTO data_store (key, json_data) VALUES (?, ?)", ('main', json.dumps(dm.db_main)))
            c.execute("INSERT OR REPLACE INTO data_store (key, json_data) VALUES (?, ?)", ('pers', json.dumps(dm.db_pers)))
            c.execute("INSERT OR REPLACE INTO data_store (key, json_data) VALUES (?, ?)", ('sos', json.dumps(dm.db_sos))) 
            c.execute("INSERT OR REPLACE INTO data_store (key, json_data) VALUES (?, ?)", ('alt', json.dumps(dm.db_alt)))
            c.execute("INSERT OR REPLACE INTO data_store (key, json_data) VALUES (?, ?)", ('pot', json.dumps(dm.db_pot)))
            c.execute("INSERT OR REPLACE INTO data_store (key, json_data) VALUES (?, ?)", ('pay', json.dumps(dm.db_pay)))
            
            for k, v in dm.preview_data.items():
                 c.execute("INSERT OR REPLACE INTO preview_store (key, content) VALUES (?, ?)", (k, v))
            conn.commit()

    def load_state(self, dm: DataManager) -> bool:
        if not os.path.exists(self.db_path): return False
        try:
            with self.get_connection() as conn:
                c = conn.cursor()
                curr = c.execute("SELECT key, value FROM settings")
                settings = dict(curr.fetchall())
                # Caselle giornaliere sempre pulite all'avvio
                dm.text_absent = ""
                dm.text_short = ""
                dm.text_classes = ""
                try: dm.total_teachers = int(settings.get('total_teachers', 0))
                except: dm.total_teachers = 0
                
                curr = c.execute("SELECT key, content FROM preview_store")
                dm.preview_data = dict(curr.fetchall()) 

                # [DB RESTORE FIX] Re-parse from raw content to ensure robustness
                if 'main' in dm.preview_data: dm.parse_main(dm.preview_data['main'], is_path=False)
                if 'pers' in dm.preview_data: dm.parse_pers(dm.preview_data['pers'], key='pers', is_path=False)
                if 'sos' in dm.preview_data: dm.parse_pers(dm.preview_data['sos'], key='sos', is_path=False)
                if 'alt' in dm.preview_data: dm.parse_alt(dm.preview_data['alt'], is_path=False)
                if 'pot' in dm.preview_data: dm.parse_generic(dm.preview_data['pot'], key='pot', is_path=False)
                if 'pay' in dm.preview_data: dm.parse_generic(dm.preview_data['pay'], key='pay', is_path=False)
                    
            return True
        except Exception as e:
            print(f"DB Load Failed: {e}")
            return False
            
    def _fix_json_load(self, json_str, fixer):
        d = json.loads(json_str)
        if isinstance(d, dict):
             for k in list(d.keys()):
                 if isinstance(d[k], dict): d[k] = fixer(d[k])
        return d

    def export_db(self, destination_path: str):
        shutil.copy2(self.db_path, destination_path)
    
    def import_db(self, source_path: str):
        shutil.copy2(source_path, self.db_path)

# --- SUBSTITUTION ENGINE ---

class SubstitutionEngine:
    def __init__(self, data_manager: DataManager):
        self.dm = data_manager
        self.pot_usage = {}  # [FIX 3.1] Contatore d'uso per round-robin potenziamento

    def find_substitutions(self, day: str, target_teachers_slugs: List[str], 
                           blacklist: List[str], short_leaves: Dict[str, List[int]], 
                           absent_classes_slugs: List[str]) -> List[Tuple[Any, ...]]:
        results = []
        freed_teachers = {h: [] for h in range(1, 9)}
        used_teachers = {}

        def cls_matches_absent(c_name: str, absent_list: List[str]) -> bool:
            if not c_name or not absent_list: return False
            c_norm = normalize_class(c_name)
            for a in absent_list:
                if not a: continue
                a_norm = normalize_class(a)
                if c_name == a or c_norm == a_norm: return True
                if len(c_name) >= 2 and len(a) >= 2 and (c_name in a or a in c_name): return True
                if len(c_norm) >= 2 and len(a_norm) >= 2 and (c_norm in a_norm or a_norm in c_norm): return True
            return False

        if absent_classes_slugs:
            for cls_key in self.dm.db_main:
                if cls_matches_absent(cls_key, absent_classes_slugs) and day in self.dm.db_main[cls_key]:
                     for h, cell in self.dm.db_main[cls_key][day].items():
                        txt = cell['text']
                        clean_txt = re.sub(r'[\-\/\.;&,]', ' ', txt.upper()); toks = [w for w in clean_txt.split() if len(w) > 1 and "COMPR" not in w and not is_subject(w)]
                        for t in toks:
                            t_slug = to_slug(t); is_absent_now = (t_slug in blacklist) or (t_slug in short_leaves and h in short_leaves[t_slug])
                            if not is_absent_now and not any(to_slug(x) == t_slug for x in freed_teachers[h]):
                                freed_teachers[h].append(t)

            for s_slug, s_data in self.dm.db_sos.items():
                if isinstance(s_data, dict) and day in s_data:
                    for h_val, sos_cls in s_data[day].items():
                        if h_val == '_name': continue
                        try: h_int = int(h_val)
                        except (ValueError, TypeError): continue
                        if sos_cls and cls_matches_absent(sos_cls, absent_classes_slugs):
                            is_absent_now = (s_slug in blacklist) or (s_slug in short_leaves and h_int in short_leaves[s_slug])
                            if not is_absent_now and not any(to_slug(x) == s_slug for x in freed_teachers[h_int]):
                                s_name = s_data.get('_name', s_slug)
                                freed_teachers[h_int].append(s_name)

        self.last_freed_teachers = {h: list(freed_teachers[h]) for h in freed_teachers}
        self.last_day = day

        def process_entry(cls, h, cell_data, source, teacher_slug):
            if teacher_slug in short_leaves:
                if h not in short_leaves[teacher_slug]: return
                
            txt = cell_data['text']; clean_txt = re.sub(r'[\-\/\.;&,]', ' ', txt.upper()); tokens = [w for w in clean_txt.split() if len(w) > 1 and "COMPR" not in w and not is_subject(w)]

            if source == 'main':
                if not tokens: return
                # [FIX] Do not return if teacher_slug is not first. We check if it is IN tokens.
                t_slug_tokens = [to_slug(x) for x in tokens]
                if teacher_slug not in t_slug_tokens: return
            
            res_txt = "CLASSE SCOPERTA"; tag = "scoperta"
            # [FIX] Subs are all tokens EXCEPT the current absent teacher (only applicable to main schedule)
            subs = [t for t in tokens if to_slug(t) != teacher_slug] if source == 'main' else []

            valid_int = []
            for s in subs:
                s_slug = to_slug(s)
                # [FIX 2.1] Solo i docenti in 'blocked' vengono bloccati, non i compresenti
                blocked = any(to_slug(b) in s_slug for b in cell_data.get('blocked', []))
                is_sub_absent = (s_slug in blacklist) or (s_slug in short_leaves and h in short_leaves[s_slug])
                if not blocked and not is_sub_absent: valid_int.append(s)

            if valid_int:
                res_txt = f"COMPRESENZA: {valid_int[0]}"
                if len(valid_int) > 1:
                    res_txt += f" (o {', '.join(valid_int[1:])})"
                tag = "compresenza"
            else:
                av_freed = [t for t in freed_teachers.get(h, []) if f"{h}-{to_slug(t)}" not in used_teachers and f"{h}-{t}" not in used_teachers]
                if av_freed:
                    sel = av_freed[0]; used_teachers[f"{h}-{to_slug(sel)}"] = True; used_teachers[f"{h}-{sel}"] = True; res_txt = f"DISP. DA CLASSE ASSENTE: {sel}"
                    if len(av_freed) > 1: res_txt += f" (altri: {', '.join(av_freed[1:])})"
                    tag = "disp_class"
                else:
                    pots = self.dm.db_pot.get(day, {}).get(h, []); valid_pots = []
                    for p in pots:
                        p_slug = to_slug(p); is_p_absent = (p_slug in blacklist) or (p_slug in short_leaves and h in short_leaves[p_slug])
                        if not is_p_absent and f"{h}-{p_slug}" not in used_teachers: valid_pots.append(p)
                    
                    if valid_pots:
                        # [FIX 3.1] Round-robin: scegli il docente con minor utilizzo nella sessione
                        sel = min(valid_pots, key=lambda p: self.pot_usage.get(to_slug(p), 0))
                        sel_slug = to_slug(sel)
                        self.pot_usage[sel_slug] = self.pot_usage.get(sel_slug, 0) + 1
                        used_teachers[f"{h}-{sel_slug}"] = True
                        res_txt = f"POTENZIAMENTO: {sel}"; tag = "potenziamento"
                    else:
                        pays = self.dm.db_pay.get(day, {}).get(h, []); valid_pays = []
                        for p in pays:
                            p_slug = to_slug(p); is_p_absent = (p_slug in blacklist) or (p_slug in short_leaves and h in short_leaves[p_slug])
                            # [FIX 3.3] Anti-sovrapposizione: verificare anche used_teachers per pagamento
                            if not is_p_absent and f"{h}-{p_slug}" not in used_teachers: valid_pays.append(p)
                        if valid_pays:
                            # Seleziona il primo disponibile e marcalo come usato
                            sel_pay = valid_pays[0]
                            used_teachers[f"{h}-{to_slug(sel_pay)}"] = True
                            res_txt = f"PAGAMENTO: {sel_pay}"
                            if len(valid_pays) > 1:
                                res_txt += f" (altri: {', '.join(valid_pays[1:])})"
                            tag = "pagamento"

            if source != 'alt':
                alt = self.dm.db_alt.get(day, {}).get(h, {}).get(cls)
                if alt:
                    a_slug = to_slug(alt); is_a_absent = (a_slug in blacklist) or (a_slug in short_leaves and h in short_leaves[a_slug])
                    if not is_a_absent:
                        alt_note = f"ALTERNATIVA: {alt} (Non avvalentesi)"
                        if tag == "scoperta":
                            res_txt = f"CLASSE SCOPERTA\n{alt_note}"
                            tag = "alternativa"
                        else:
                            res_txt += f"\n{alt_note}"
            else:
                if tag == "scoperta":
                    res_txt = "CLASSE SCOPERTA (Non avvalentesi)"
                else:
                    res_txt += " (Non avvalentesi)"

            # [FIX] Return teacher_slug as 6th element for stats
            results.append((cls, h, txt, res_txt, tag, teacher_slug))

        # Fase 1: orario principale, personale, sostegno
        for teacher_slug in target_teachers_slugs:
            for cls in self.dm.db_main:
                if cls_matches_absent(cls, absent_classes_slugs): continue
                if day in self.dm.db_main[cls]:
                    for h in range(1, 9):
                        if h in self.dm.db_main[cls][day]:
                            cell = self.dm.db_main[cls][day][h]
                            clean_txt = re.sub(r'[\-\/\.;&,]', ' ', cell['text'].upper())
                            tokens = [to_slug(w) for w in clean_txt.split() if len(w) > 1 and "COMPR" not in w and not is_subject(w)]
                            if teacher_slug in tokens:
                                existing_idx = next((i for i, r in enumerate(results) if r[0] == cls and r[1] == h), None)
                                if existing_idx is not None and not cell.get('blocked'):
                                    prev_r = results[existing_idx]
                                    prev_slugs = str(prev_r[5]).split(',')
                                    if teacher_slug not in prev_slugs:
                                        combined_slugs = f"{prev_r[5]},{teacher_slug}"
                                        results[existing_idx] = (prev_r[0], prev_r[1], prev_r[2], prev_r[3], prev_r[4], combined_slugs)
                                    continue
                                process_entry(cls, h, cell, 'main', teacher_slug)
            
            if teacher_slug in self.dm.db_pers and day in self.dm.db_pers[teacher_slug]:
                 for h, cls in self.dm.db_pers[teacher_slug][day].items():
                    if cls_matches_absent(cls, absent_classes_slugs): continue
                    existing_idx = next((i for i, r in enumerate(results) if r[0] == cls and r[1] == h), None)
                    if existing_idx is not None:
                        prev_r = results[existing_idx]
                        prev_slugs = str(prev_r[5]).split(',')
                        if teacher_slug not in prev_slugs:
                            combined_slugs = f"{prev_r[5]},{teacher_slug}"
                            results[existing_idx] = (prev_r[0], prev_r[1], prev_r[2], prev_r[3], prev_r[4], combined_slugs)
                        continue
                    process_entry(cls, h, {'text': f'(Da Personale: {teacher_slug})', 'blocked': []}, 'pers', teacher_slug)
            
            if teacher_slug in self.dm.db_sos and day in self.dm.db_sos[teacher_slug]:
                 for h, cls in self.dm.db_sos[teacher_slug][day].items():
                    if cls_matches_absent(cls, absent_classes_slugs): continue
                    process_entry(cls, h, {'text': f'(Da Sostegno: {teacher_slug})', 'blocked': []}, 'pers', teacher_slug)

        # Fase 2: orario alternativa
        # Se il docente di alternativa è assente ma il docente di religione è presente,
        # individua il docente in sostituzione per gli alunni non avvalentesi.
        # Se mancano entrambi, la classe è già coperta dal sostituto di religione sopra (un solo sostituto).
        if day in self.dm.db_alt:
            for teacher_slug in target_teachers_slugs:
                for h in range(1, 9):
                    if h in self.dm.db_alt[day]:
                        for cls, alt_teacher in self.dm.db_alt[day][h].items():
                            if to_slug(alt_teacher) == teacher_slug:
                                if cls_matches_absent(cls, absent_classes_slugs): continue
                                if teacher_slug in short_leaves and h not in short_leaves[teacher_slug]: continue
                                # Se per questa classe e ora è già stata individuata una sostituzione, salta (un solo sostituto)
                                if any(r[0] == cls and r[1] == h for r in results):
                                    continue
                                rel_note = " (Religione presente)" if (cls in self.dm.db_main and day in self.dm.db_main[cls] and h in self.dm.db_main[cls][day]) else ""
                                cell_info = {'text': f'MAT. ALTERNATIVA: {alt_teacher}{rel_note}', 'blocked': []}
                                process_entry(cls, h, cell_info, 'alt', teacher_slug)

        return results

# --- OPTIMIZED EDITOR GUI ---

# ── Palette colori editor ──────────────────────────────────────────────
_EDITOR_COLORS = {
    'toolbar_bg':     '#1E293B',   # Slate scuro per la toolbar
    'toolbar_fg':     '#F1F5F9',   # Testo chiaro sulla toolbar
    'toolbar_accent': '#38BDF8',   # Azzurro cielo per accenti
    'btn_save_bg':    '#22C55E',   # Verde vivace per Salva
    'btn_save_fg':    '#FFFFFF',
    'btn_save_hover': '#16A34A',
    'heading_bg':     '#334155',   # Slate medio per le intestazioni colonne
    'heading_fg':     '#F8FAFC',
    'row_even':       '#FFFFFF',   # Bianco per righe pari
    'row_odd':        '#F1F5F9',   # Grigio ghiaccio per righe dispari
    'row_selected':   '#DBEAFE',   # Azzurro tenue per selezione
    'row_sel_fg':     '#1E3A5F',
    'cell_border':    '#CBD5E1',   # Bordo sottile tra celle
    'status_bg':      '#F8FAFC',   # Sfondo barra di stato
    'status_fg':      '#64748B',   # Testo grigio nella status bar
    'overflow_bg':    '#FEF3C7',   # Giallo caldo per avviso overflow
    'overflow_fg':    '#92400E',
    'edit_bg':        '#FFFBEB',   # Sfondo cella in editing
    'edit_border':    '#F59E0B',   # Bordo cella in editing
    'window_bg':      '#F1F5F9',   # Sfondo finestra
}


class EditorWindow(tk.Toplevel):
    """Editor CSV ottimizzato con grafica premium e protezione overflow righe."""

    def __init__(self, parent, title, content, on_save, max_rows=1000):
        super().__init__(parent)
        self.title(f"📊 Editor — {title}")
        self.geometry("1400x800")
        self.configure(bg=_EDITOR_COLORS['window_bg'])
        self.on_save = on_save
        self.max_rows = max_rows
        self.overflow_rows = []  # [FIX 1.1] Righe oltre il limite visualizzato, preservate al salvataggio

        # ── Stile Treeview (una sola volta, leggero) ───────────────────
        self._apply_style()

        # ── Toolbar moderna ────────────────────────────────────────────
        toolbar = tk.Frame(self, bg=_EDITOR_COLORS['toolbar_bg'], padx=12, pady=8)
        toolbar.pack(fill="x")

        # Pulsante Salva
        self.btn_save = tk.Button(
            toolbar, text="💾  SALVA E CHIUDI",
            bg=_EDITOR_COLORS['btn_save_bg'], fg=_EDITOR_COLORS['btn_save_fg'],
            activebackground=_EDITOR_COLORS['btn_save_hover'],
            font=("Helvetica", 11, "bold"), relief="flat",
            padx=16, pady=4, cursor="hand2", command=self.save,
            highlightbackground=_EDITOR_COLORS['btn_save_bg']
        )
        self.btn_save.pack(side="left")

        # Hint editing
        hint_frame = tk.Frame(toolbar, bg=_EDITOR_COLORS['toolbar_bg'])
        hint_frame.pack(side="left", padx=20)
        tk.Label(hint_frame, text="✏️", font=("Segoe UI", 12),
                 bg=_EDITOR_COLORS['toolbar_bg']).pack(side="left")
        tk.Label(hint_frame,
                 text="Doppio click su una cella per modificare  •  Invio per confermare  •  Esc per annullare",
                 bg=_EDITOR_COLORS['toolbar_bg'], fg=_EDITOR_COLORS['toolbar_accent'],
                 font=("Segoe UI", 10)).pack(side="left", padx=6)

        # ── Avviso overflow (se presente, verrà mostrato dopo il parsing) ──
        self._overflow_frame = tk.Frame(toolbar, bg=_EDITOR_COLORS['toolbar_bg'])
        self._overflow_frame.pack(side="right")

        # Parse content
        self.delimiter = ';'
        self.rows = []
        if content:
            first_line = content.splitlines()[0] if content.splitlines() else ""
            if first_line.count(';') >= first_line.count(','): self.delimiter = ';'
            else: self.delimiter = ','
            try: 
                reader = csv.reader(content.splitlines(), delimiter=self.delimiter)
                self.rows = list(reader)
            except: self.rows = []
        
        if not self.rows: self.rows = [["" for _ in range(60)] for _ in range(100)]
        
        # Treeview setup
        max_cols = 0
        for r in self.rows: max_cols = max(max_cols, len(r))
        max_cols = max(max_cols, 20) # At least 20
        max_cols = min(max_cols, 100) # Cap at 100

        # Contenitore con bordo sottile
        tree_container = tk.Frame(self, bg=_EDITOR_COLORS['cell_border'], bd=1, relief="solid")
        tree_container.pack(fill="both", expand=True, padx=8, pady=(4, 0))
        
        self.tree = ttk.Treeview(tree_container, columns=[str(i) for i in range(max_cols)],
                                  show="headings", style="Editor.Treeview")
        
        # Scrollbars
        vsb = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True)
        
        # Intestazioni e colonne
        for i in range(max_cols):
            self.tree.heading(str(i), text=f"  Col {i+1}  ")
            self.tree.column(str(i), width=180, minwidth=80, stretch=False)
        
        # ── Zebra-striping con tag (applicati durante l'insert, costo zero) ──
        self.tree.tag_configure("row_even", background=_EDITOR_COLORS['row_even'])
        self.tree.tag_configure("row_odd", background=_EDITOR_COLORS['row_odd'])

        for i, row in enumerate(self.rows[:self.max_rows]):
            vals = row + [""] * (max_cols - len(row))
            tag = "row_even" if i % 2 == 0 else "row_odd"
            self.tree.insert("", "end", iid=str(i), values=vals, tags=(tag,))
        
        # [FIX 1.1] Conserva righe non visualizzate per evitare troncamento al salvataggio
        if len(self.rows) > self.max_rows:
            self.overflow_rows = self.rows[self.max_rows:]
            overflow_pill = tk.Frame(self._overflow_frame, bg=_EDITOR_COLORS['overflow_bg'],
                                     padx=10, pady=2, highlightbackground=_EDITOR_COLORS['overflow_fg'],
                                     highlightthickness=1)
            overflow_pill.pack(side="right")
            tk.Label(overflow_pill,
                     text=f"⚠ {len(self.overflow_rows)} righe nascoste (preservate al salvataggio)",
                     bg=_EDITOR_COLORS['overflow_bg'], fg=_EDITOR_COLORS['overflow_fg'],
                     font=("Segoe UI", 9, "bold")).pack()

        # ── Barra di stato ─────────────────────────────────────────────
        status_bar = tk.Frame(self, bg=_EDITOR_COLORS['status_bg'], padx=12, pady=4)
        status_bar.pack(fill="x", side="bottom")

        rows_shown = min(len(self.rows), self.max_rows)
        total_rows = len(self.rows)
        self.lbl_status = tk.Label(
            status_bar,
            text=f"📄 {rows_shown} righe visualizzate  •  {max_cols} colonne  •  {total_rows} righe totali  •  Delimitatore: «{self.delimiter}»",
            bg=_EDITOR_COLORS['status_bg'], fg=_EDITOR_COLORS['status_fg'],
            font=("Segoe UI", 9)
        )
        self.lbl_status.pack(side="left")

        self.lbl_file_title = tk.Label(
            status_bar, text=title,
            bg=_EDITOR_COLORS['status_bg'], fg=_EDITOR_COLORS['heading_bg'],
            font=("Segoe UI", 9, "bold")
        )
        self.lbl_file_title.pack(side="right")

        self.tree.bind("<Double-1>", self.on_double_click)

    def _apply_style(self):
        """Applica lo stile personalizzato al Treeview dell'editor."""
        style = ttk.Style(self)

        style.configure("Editor.Treeview",
                         background=_EDITOR_COLORS['row_even'],
                         foreground="#1E293B",
                         rowheight=28,
                         font=("Segoe UI", 10),
                         borderwidth=0,
                         fieldbackground=_EDITOR_COLORS['row_even'])

        style.configure("Editor.Treeview.Heading",
                         background=_EDITOR_COLORS['heading_bg'],
                         foreground=_EDITOR_COLORS['heading_fg'],
                         font=("Segoe UI", 10, "bold"),
                         relief="flat",
                         padding=(6, 4))

        style.map("Editor.Treeview",
                   background=[("selected", _EDITOR_COLORS['row_selected'])],
                   foreground=[("selected", _EDITOR_COLORS['row_sel_fg'])])

        style.map("Editor.Treeview.Heading",
                   background=[("active", "#475569")])

        # Layout senza bordi interni
        style.layout("Editor.Treeview", [('Treeview.treearea', {'sticky': 'nswe'})])

    def on_double_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell": return
        
        col_id = self.tree.identify_column(event.x) # #1, #2..
        row_id = self.tree.identify_row(event.y)
        if not col_id or not row_id: return
        
        col_idx = int(col_id.replace("#", "")) - 1
        
        # Get bounds
        try:
            x, y, w, h = self.tree.bbox(row_id, col_id)
        except (ValueError, tk.TclError):
            return
        
        # Current val
        vals = list(self.tree.item(row_id, "values"))
        curr_val = vals[col_idx] if col_idx < len(vals) else ""
        
        # Create styled entry
        entry = tk.Entry(self.tree, width=20,
                         font=("Segoe UI", 10),
                         bg=_EDITOR_COLORS['edit_bg'],
                         fg="#1E293B",
                         insertbackground="#1E293B",
                         highlightcolor=_EDITOR_COLORS['edit_border'],
                         highlightbackground=_EDITOR_COLORS['edit_border'],
                         highlightthickness=2,
                         relief="flat",
                         bd=0)
        entry.place(x=x, y=y, width=max(w, 120), height=h)
        entry.insert(0, curr_val)
        entry.select_range(0, "end")
        entry.focus_set()
        
        def commit(e=None):
            new_val = entry.get()
            vals = list(self.tree.item(row_id, "values"))
            while len(vals) <= col_idx: vals.append("")
            vals[col_idx] = new_val
            self.tree.item(row_id, values=vals)
            entry.destroy()
        
        def cancel(e=None):
            entry.destroy()
            
        entry.bind("<Return>", commit)
        entry.bind("<FocusOut>", commit)
        entry.bind("<Escape>", cancel)

    def save(self):
        output = io.StringIO()
        try:
            # Determine correct quoting dialect or force one
            writer = csv.writer(output, delimiter=self.delimiter, quoting=csv.QUOTE_MINIMAL)
            
            for child in self.tree.get_children():
                vals = self.tree.item(child, "values")
                # Convert tuple to list
                row = list(vals)
                # clean trailing empty
                while row and not row[-1]: row.pop()
                if any(row):
                    writer.writerow([str(x) for x in row])
            
            # [FIX 1.1] Riaccodare le righe in overflow che non erano visualizzate
            for row in self.overflow_rows:
                # clean trailing empty
                row_copy = list(row)
                while row_copy and not row_copy[-1]: row_copy.pop()
                if any(row_copy):
                    writer.writerow([str(x) for x in row_copy])
            
            new_content = output.getvalue()
            self.on_save(new_content)
            self.destroy()
        except Exception as e:
            messagebox.showerror("Errore Salvataggio", f"Impossibile salvare il contenuto CSV:\n{e}")



class App(tk.Tk):
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
        self.loaded_paths = {} # [PERSISTENCE] Track file paths
        self.temp_files = []   # [FIX 4.7] Lista file temporanei da pulire alla chiusura
        self.setup_ui()
        self.setup_menu()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def setup_icons(self):
        ico_path = Configuration.resource_path("Alberghiero.ico")
        png_path = Configuration.resource_path("alberghiero.png")
        if os.path.exists(png_path):
            try: self.iconphoto(True, tk.PhotoImage(file=png_path))
            except: pass
        elif os.path.exists(ico_path):
            try: self.iconbitmap(ico_path)
            except: pass
        self.logo_path = png_path
        
    def setup_menu(self):
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
            manual_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "manuale_utente.html")
        
        if os.path.exists(manual_path):
            try:
                webbrowser.open('file://' + os.path.abspath(manual_path))
            except Exception as e:
                messagebox.showerror("Errore", f"Impossibile aprire il manuale:\n{e}")
        else:
            messagebox.showerror("Errore", "File manuale_utente.html non trovato.")

    def show_guide(self):
        top = tk.Toplevel(self)
        top.title("Guida Utilizzo")
        top.geometry("640x650")
        top.configure(bg="#FAFAFA")
        
        lbl = tk.Label(top, text="MANUALE UTENTE — GUIDA RAPIDA", font=("Helvetica", 16, "bold"), bg="#FAFAFA", fg="#1565C0")
        lbl.pack(pady=(12, 4))
        
        btn_open_html = tk.Button(top, text="🌐 Apri Manuale Completo in HTML", font=("Segoe UI", 10, "bold"),
                                  bg="#E3F2FD", fg="#1565C0", relief="groove", padx=12, pady=4, cursor="hand2",
                                  command=self.open_manual_html)
        btn_open_html.pack(pady=(0, 8))
        
        txt_frame = tk.Frame(top, bg="#FAFAFA")
        txt_frame.pack(fill="both", expand=True, padx=15, pady=(0, 10))
        
        txt = tk.Text(txt_frame, wrap="word", bg="#FFFFFF", font=("Segoe UI", 10), bd=1, relief="solid", padx=15, pady=10)
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
            ("Selezionare il giorno della settimana e cliccare sul pulsante verde 'PROCEDI'.\nIl programma applica la gerarchia: Compresenza → Disp. da classe assente → Potenziamento (round-robin) → Ore a pagamento → Classe scoperta.", "body"),
            
            ("\n4. Menu Contestuale (Tasto Destro)", "h1"),
            ("Facendo clic con il tasto destro su qualsiasi riga della tabella è possibile:\n- Segnare la riga come 'Coperta manualmente' (sfondo verde salvia);\n- Modificare manualmente il testo del sostituto;\n- Rimuovere la riga dalla tabella;\n- Copiare i dati negli appunti.", "body"),
            
            ("\n5. Esportazione e Stampa", "h1"),
            ("• PDF (Blu): Anteprima di stampa con layout ufficiale dell'Istituto.\n• XLSX (Verde): Esportazione foglio di calcolo Excel.\n• WORD (Rosso): Esportazione DOCX con scelta tra stile 'A Colori' (con evidenziazioni fedeli) o 'Bianco e Nero' per stampa classica.", "body"),
            
            ("\n⭐ Novità della Versione 3.0", "h2"),
            ("• Esportazione Word A Colori / B-N con selezione grafica personalizzata.", "bullet"),
            ("• Gestione Materia Alternativa corretta: sostituzione per gli alunni non avvalentesi quando manca solo l'alternativa (con religione presente), e singolo docente sostituto se mancano entrambi.", "bullet"),
            ("• Grafica tabella ad alta leggibilità (rowheight 52px e colonna espandibile): nessun testo tagliato.", "bullet"),
            ("• Avvio ripulito: le caselle giornaliere partono sempre pulite e pronte per il nuovo giorno.", "bullet"),
            ("• Algoritmo di potenziamento round-robin per un'equa rotazione dei docenti.", "bullet"),
            ("• Editor CSV anti-troncamento: salvataggio sicuro e zero perdita di righe.", "bullet"),
            ("• Menu contestuale tasto destro completo per gestione rapida delle righe.", "bullet")
        ]
        
        for text, tag in content:
            txt.insert("end", text + "\n", tag)
            
        txt.config(state="disabled")

    def show_about(self):
        top = tk.Toplevel(self)
        top.title("About")
        top.geometry("420x330")
        top.configure(bg="white")
        top.resizable(False, False)
        
        tk.Label(top, text="Gestione Sostituzioni", font=("Helvetica", 22, "bold"), bg="white", fg="#2c3e50").pack(pady=(30, 8))
        tk.Label(top, text="Versione 3.3", font=("Segoe UI", 13, "bold"), bg="white", fg="#1976D2").pack(pady=2)
        tk.Label(top, text="Ultimo aggiornamento: 31 Ottobre 2026", font=("Segoe UI", 11), bg="white", fg="#555555").pack(pady=4)
        tk.Label(top, text="Ultimo aggiornamento: 24 Ottobre 2026", font=("Segoe UI", 11), bg="white", fg="#555555").pack(pady=4)
        tk.Label(top, text="Mario Carta copyright © 2026", font=("Segoe UI", 11, "italic"), bg="white", fg="#2c3e50").pack(pady=(25, 15))
        tk.Button(top, text="Chiudi", font=("Segoe UI", 10), bg="#F5F5F5", relief="groove", padx=15, pady=3, command=top.destroy).pack(pady=10)

    def reset_all_data(self):
        if not messagebox.askyesno("Attenzione", "Questa operazione è irreversibile!\n\nTutti i dati caricati, le statistiche e il database verranno cancellati completamente.\n\nVuoi davvero procedere?"):
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
        self.entry_total_teachers.delete(0, "end"); self.entry_total_teachers.insert(0, "0")
        
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
            print(f"Error wiping DB: {e}")
            
        # 5. Refresh UI
        self.refresh_labels()
        self.update_stats_display()
        
        messagebox.showinfo("Reset Completato", "Tutti i dati sono stati cancellati e il programma è stato azzerato.")

    def setup_ui(self):
        style = ttk.Style()
        style.theme_use('clam')
        bg = Configuration.THEME['bg']; primary = Configuration.THEME['primary']
        
        style.configure("Treeview", background="#FFFFFF", foreground="black", rowheight=52, font=("Segoe UI", 11), borderwidth=1)
        style.configure("Treeview.Heading", font=("Segoe UI", 12, "bold"), background=primary, foreground="white", relief="flat")
        style.layout("Treeview", [('Treeview.treearea', {'sticky': 'nswe'})]) 
        
        # Header
        header_frame = tk.Frame(self, bg=primary); header_frame.pack(fill="x")
        h_inner = tk.Frame(header_frame, bg=primary); h_inner.pack(pady=5, padx=30, anchor="w")
        if self.logo_path and os.path.exists(self.logo_path):
            try:
                self.img_header = tk.PhotoImage(file=self.logo_path).subsample(6, 6) 
                tk.Label(h_inner, image=self.img_header, bg=primary).pack(side="left", padx=(0, 15))
            except: pass
        tk.Label(h_inner, text="GESTIONE SOSTITUZIONI", font=("Helvetica", 24, "bold"), fg="white", bg=primary).pack(side="left")
        
        # Copyright on the right of the title
        tk.Label(h_inner, text="© Mario Carta", font=("Helvetica", 10), fg="white", bg=primary).pack(side="left", padx=(10, 0), anchor="cen")
        
        # STATS FRAME
        stats_frame = tk.Frame(self, bg="#E3F2FD", bd=1, relief="solid", padx=20, pady=2)
        stats_frame.pack(fill="x", padx=30, pady=2)
        
        tk.Label(stats_frame, text="STATISTICHE ASSENZE", font=("Segoe UI", 11, "bold"), bg="#E3F2FD", fg="#1565C0").pack(side="left")
        
        tk.Label(stats_frame, text="Totale Organico:", bg="#E3F2FD", fg="#000000", font=("Segoe UI", 10)).pack(side="left", padx=(20, 5))
        self.entry_total_teachers = tk.Entry(stats_frame, width=5, bg="white", fg="black", insertbackground="black")
        self.entry_total_teachers.pack(side="left")
        self.entry_total_teachers.insert(0, str(self.dm.total_teachers))
        
        self.lbl_stats = tk.Label(stats_frame, text="Caricamento...", bg="#E3F2FD", fg="#000000", font=("Segoe UI", 10))
        self.lbl_stats.pack(side="left", padx=20)
        
        tk.Button(stats_frame, text="Resetta Statistiche", bg="#ffcdd2", command=self.reset_stats, font=("Segoe UI", 8)).pack(side="right", padx=5)
        tk.Button(stats_frame, text="Stampa Statistiche", bg="#FFF", command=self.print_stats, font=("Segoe UI", 8)).pack(side="right", padx=5)
        
        self.update_stats_display()

        main_frame = tk.Frame(self, bg=bg, padx=10, pady=2); main_frame.pack(fill="both", expand=True)

        # GRID FILES
        grid_frame = tk.Frame(main_frame, bg=bg); grid_frame.pack(fill="x", pady=2)
        self.create_file_box(grid_frame, "1. ORARIO CLASSI", "main", 0, 0, "#00695C", lambda p: self.dm.parse_main(p))
        self.create_file_box(grid_frame, "2. ORARIO PERSONALE", "pers", 0, 1, "#0277BD", lambda p: self.dm.parse_pers(p, 'pers'))
        self.create_file_box(grid_frame, "3. MATERIA ALTERNATIVA", "alt", 0, 2, "#6A1B9A", lambda p: self.dm.parse_alt(p))
        self.create_file_box(grid_frame, "4. POTENZIAMENTO", "pot", 1, 0, "#EF6C00", lambda p: self.dm.parse_generic(p, 'pot'))
        self.create_file_box(grid_frame, "5. A PAGAMENTO", "pay", 1, 1, "#2E7D32", lambda p: self.dm.parse_generic(p, 'pay'))
        self.create_file_box(grid_frame, "6. DOCENTI SOSTEGNO", "sos", 1, 2, "#d35400", lambda p: self.dm.parse_pers(p, 'sos'))

        # TEXT AREAS
        lists_frame = tk.Frame(main_frame, bg=bg); lists_frame.pack(fill="x", pady=2)
        lists_frame.columnconfigure(0, weight=1); lists_frame.columnconfigure(1, weight=1); lists_frame.columnconfigure(2, weight=1)
        
        self.txt_absent = self.create_styled_text_box(lists_frame, "DOCENTI ASSENTI", 0, 0, "#c0392b", bg_color="#FFEBEE") # Dark Red border
        
        self.txt_short_leave = self.create_styled_text_box(lists_frame, "PERMESSI BREVI", 0, 1, "#16a085", "Es: Rossi 1 2", bg_color="#E0F7FA") # Green/Teal border
        
        self.txt_classes_absent = self.create_styled_text_box(lists_frame, "CLASSI ASSENTI", 0, 2, "#f39c12", bg_color="#FFFDE7") # Orange border

        # CONTROLS
        ctrl_frame = tk.Frame(main_frame, bg=bg); ctrl_frame.pack(fill="x", pady=8)
        
        # Left side control group
        c_left = tk.Frame(ctrl_frame, bg=bg); c_left.pack(side="left", fill="x", expand=True)
        
        tk.Label(c_left, text="GIORNO:", font=("Segoe UI", 12, "bold"), bg=bg, fg=Configuration.THEME['text']).pack(side="left")
        self.combo_day = ttk.Combobox(c_left, values=Configuration.DAYS[1:], state="readonly", width=12, font=("Segoe UI", 12))
        self.combo_day.current(0); self.combo_day.pack(side="left", padx=10)
        
        tk.Label(c_left, text="DOCENTI DA SOSTITUIRE:", font=("Segoe UI", 12, "bold"), bg=bg, fg=Configuration.THEME['text']).pack(side="left", padx=(15, 5))
        self.entry_teacher = tk.Entry(c_left, font=("Segoe UI", 13), relief="flat", bd=2, highlightthickness=1, highlightbackground="#ccc", bg="white", fg="black", insertbackground="black")
        self.entry_teacher.pack(side="left", fill="x", expand=True, padx=5)
        
        # Right side buttons
        btn_frame = tk.Frame(ctrl_frame, bg=bg); btn_frame.pack(side="right", padx=5)
        self.create_btn(btn_frame, "PROCEDI", "#4CAF50", self.run_calculation, fg="white")
        self.create_btn(btn_frame, "SVUOTA", "#f44336", self.clear_table, fg="white")
        
        # Export Bar
        export_frame = tk.Frame(main_frame, bg=bg); export_frame.pack(fill="x", pady=2)
        self.create_btn(export_frame, "ANTEPRIMA STAMPA (PDF)", "#1565C0", self.print_preview, side="left", fg="white") 
        if HAS_PANDAS: self.create_btn(export_frame, "XLSX", "#2E7D32", self.export_excel, side="left", fg="white")
        if HAS_DOCX: self.create_btn(export_frame, "WORD", "#c62828", self.export_word, side="left", fg="white")
        
        self.create_btn(export_frame, "IMPORTA DB", "#FFFFFF", self.import_db_file, side="right")
        self.create_btn(export_frame, "SALVA DB", "#FFFFFF", self.export_db_file, side="right")

        tree_frame = tk.Frame(main_frame, bg="gray", bd=1, relief="solid") 
        tree_frame.pack(fill="both", expand=True, pady=5)
        self.tree = ttk.Treeview(tree_frame, columns=("class", "hour", "orig", "sub"), show="headings", height=15)
        self.tree.heading("class", text="CLASSE"); self.tree.column("class", width=90, minwidth=70, anchor="center")
        self.tree.heading("hour", text="ORA"); self.tree.column("hour", width=60, minwidth=50, anchor="center")
        self.tree.heading("orig", text="SITUAZIONE"); self.tree.column("orig", width=320, minwidth=250, anchor="w")
        self.tree.heading("sub", text="SOSTITUZIONE"); self.tree.column("sub", width=650, minwidth=400, anchor="w", stretch=True)
        
        sb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=sb.set); sb.pack(side="right", fill="y"); self.tree.pack(side="left", fill="both", expand=True)
        for k, col in Configuration.COLORS.items(): self.tree.tag_configure(k, background=col)
        # [FIX 4.1] Tag aggiuntivo per righe marcate come coperte
        self.tree.tag_configure('coperta', background='#A5D6A7')
        
        # [FIX 4.1] Menu contestuale (tasto destro)
        self.setup_context_menu()
        
        self.refresh_labels()

    # --- [FIX 4.1] MENU CONTESTUALE (TASTO DESTRO) ---
    
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
        if not selected: return
        for item in selected:
            vals = list(self.tree.item(item, "values"))
            vals[3] = "COPERTA MANUALMENTE"
            self.tree.item(item, values=vals, tags=("coperta",))
    
    def edit_substitute(self):
        """Permette di selezionare o modificare il sostituto per la riga selezionata."""
        selected = self.tree.selection()
        if not selected: return
        item = selected[0]
        vals = list(self.tree.item(item, "values"))
        tags = list(self.tree.item(item, "tags"))
        if len(vals) < 4: return

        cls_val = str(vals[0]); hour_val = str(vals[1]); sit_val = str(vals[2]); current_sub = str(vals[3])
        try: h_int = int(re.search(r'\d+', hour_val).group(0))
        except Exception: h_int = 1
        day = self.combo_day.get()

        candidates = []
        m_disp = re.match(r'^DISP\.\s*DA\s*CLASSE\s*ASSENTE:\s*([^\(]+)(?:\((?:altri|o):\s*([^\)]+)\))?', current_sub, re.IGNORECASE)
        if m_disp:
            main_t = m_disp.group(1).strip()
            if main_t: candidates.append(f"DISP. DA CLASSE ASSENTE: {main_t}")
            if m_disp.group(2):
                for extra in m_disp.group(2).split(','):
                    extra = extra.strip()
                    if extra: candidates.append(f"DISP. DA CLASSE ASSENTE: {extra}")

        if hasattr(self, 'engine') and hasattr(self.engine, 'last_freed_teachers'):
            for ft in self.engine.last_freed_teachers.get(h_int, []):
                opt = f"DISP. DA CLASSE ASSENTE: {ft}"
                if opt not in candidates: candidates.append(opt)

        if hasattr(self, 'dm') and self.dm:
            for s_slug, s_data in self.dm.db_sos.items():
                if isinstance(s_data, dict) and day in s_data and h_int in s_data[day]:
                    sos_cls = s_data[day][h_int]
                    if sos_cls == cls_val or (len(sos_cls) >= 3 and (sos_cls in cls_val or cls_val in sos_cls)):
                        s_name = s_data.get('_name', s_slug)
                        opt = f"COMPRESENZA SOSTEGNO: {s_name}"
                        if opt not in candidates: candidates.append(opt)

            for p in self.dm.db_pot.get(day, {}).get(h_int, []):
                opt = f"POTENZIAMENTO: {p}"
                if opt not in candidates: candidates.append(opt)

            for p in self.dm.db_pay.get(day, {}).get(h_int, []):
                opt = f"PAGAMENTO: {p}"
                if opt not in candidates: candidates.append(opt)

        candidates.append("COPERTA MANUALMENTE")
        candidates.append("CLASSE SCOPERTA")

        clean_candidates = []
        seen = set()
        for c in candidates:
            if c not in seen: seen.add(c); clean_candidates.append(c)

        dialog = tk.Toplevel(self)
        dialog.title(f"Scegli Sostituto — Classe {cls_val} ({hour_val}ª Ora)")
        dialog.geometry("560x320"); dialog.configure(bg="#F8FAFC"); dialog.transient(self); dialog.grab_set(); dialog.resizable(False, False)

        hdr_frame = tk.Frame(dialog, bg="#1E293B", pady=10, padx=16); hdr_frame.pack(fill="x")
        tk.Label(hdr_frame, text=f"Classe: {cls_val}   |   Ora: {hour_val}ª", font=("Segoe UI", 12, "bold"), fg="#38BDF8", bg="#1E293B").pack(anchor="w")
        tk.Label(hdr_frame, text=f"Situazione: {sit_val}", font=("Segoe UI", 10), fg="#CBD5E1", bg="#1E293B").pack(anchor="w", pady=(2, 0))

        btn_frame = tk.Frame(dialog, bg="#F8FAFC", padx=20); btn_frame.pack(side="bottom", fill="x", pady=(0, 15))

        content_frame = tk.Frame(dialog, bg="#F8FAFC", padx=20, pady=12); content_frame.pack(fill="both", expand=True)
        tk.Label(content_frame, text="Seleziona dalla lista dei docenti disponibili a quest'ora:\n(oppure digita un messaggio personalizzato, es: 'La classe entra alle 9:15')",
                 font=("Segoe UI", 10, "bold"), bg="#F8FAFC", fg="#334155", justify="left").pack(anchor="w", pady=(0, 8))

        cmb_sub = ttk.Combobox(content_frame, values=clean_candidates, font=("Segoe UI", 11), width=48)
        cmb_sub.set(current_sub); cmb_sub.pack(fill="x", pady=5); cmb_sub.focus()

        def apply_selection():
            new_val = cmb_sub.get().strip()
            if new_val:
                vals[3] = new_val
                new_tag = tags[0] if tags else "scoperta"
                if "COPERTA MANUALMENTE" in new_val: new_tag = "coperta"
                elif "DISP. DA CLASSE ASSENTE" in new_val: new_tag = "disp_class"
                elif "SOSTEGNO" in new_val: new_tag = "sostegno"
                elif "COMPRESENZA" in new_val: new_tag = "compresenza"
                elif "POTENZIAMENTO" in new_val: new_tag = "potenziamento"
                elif "PAGAMENTO" in new_val: new_tag = "pagamento"
                elif "CLASSE SCOPERTA" in new_val: new_tag = "scoperta"
                self.tree.item(item, values=vals, tags=(new_tag,))
            dialog.destroy()

        def make_btn(parent, text, bg_color, command):
            lbl = tk.Label(parent, text=text, font=("Segoe UI", 10, "bold"), bg=bg_color, fg="white", padx=16, pady=6, cursor="hand2", relief="raised", bd=1)
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
        if not selected: return
        for item in selected:
            self.tree.delete(item)
    
    def copy_row(self):
        """Copia il contenuto della riga selezionata negli appunti."""
        selected = self.tree.selection()
        if not selected: return
        item = selected[0]
        vals = self.tree.item(item, "values")
        text = "\t".join(str(v) for v in vals)
        self.clipboard_clear()
        self.clipboard_append(text)

    def create_btn(self, parent, text, bg, cmd, side="right", fg="black"):
        # [MAC FIX] Tkinter buttons on Mac ignore 'bg' for the face, keeping it white/gray.
        # White text on white face is invisible. Force black text on Mac.
        if sys.platform == 'darwin':
            fg = "black"
        
        tk.Button(parent, text=text, bg=bg, activebackground="#E0E0E0", fg=fg, highlightbackground=bg,
                  font=("Segoe UI", 9, "bold"),
                  command=cmd, padx=10, pady=3, relief="raised", borderwidth=2, cursor="hand2").pack(side=side, padx=3)

    def create_file_box(self, parent, title, key, r, c, color, load_callback):
        f = tk.Frame(parent, bg="white", highlightbackground=color, highlightthickness=1, padx=5, pady=2)
        f.grid(row=r, column=c, sticky="ew", padx=5, pady=2)
        parent.columnconfigure(c, weight=1)
        tk.Label(f, text=title, bg="white", font=("Segoe UI", 10, "bold"), fg="#2c3e50").pack(anchor="w")
        lbl = tk.Label(f, text="Nessun dato", bg="white", fg="#888", font=("Segoe UI", 9))
        lbl.pack(anchor="w", pady=0)
        
        def on_click():
            path = filedialog.askopenfilename(filetypes=[("CSV", "*.csv")])
            if path:
                n = load_callback(path); lbl.config(text=f"Caricati: {n}", fg="#2E7D32")
                self.loaded_paths[key] = path # [PERSISTENCE] Store path
                
        def on_preview():
            self.show_editor(key, title)
            
        tk.Frame(f, height=1, bg="#eee").pack(fill="x", pady=2)
        btn_box = tk.Frame(f, bg="white")
        btn_box.pack(fill="x")
        tk.Button(btn_box, text="👁 Anteprima", bg="#FAFAFA", fg="#333", relief="flat", command=on_preview, font=("Segoe UI", 9)).pack(side="left")
        tk.Button(btn_box, text="📂 Carica", bg="#F5F5F5", fg="black", relief="flat", command=on_click, font=("Segoe UI", 9)).pack(side="right")
        self.file_boxes[key] = {'label': lbl, 'loader': load_callback}
        return f
    
    def show_editor(self, key, title):
        content = self.dm.preview_data.get(key, "")
        limit = 400
        if key == 'main': limit = 600
        elif key == 'pers': limit = 1000
        elif key == 'sos': limit = 700

        def save_callback(new_content):
            try:
                if key == 'main': self.dm.parse_main(new_content, is_path=False)
                elif key == 'pers' or key == 'sos': self.dm.parse_pers(new_content, key=key, is_path=False)
                elif key == 'alt': self.dm.parse_alt(new_content, is_path=False)
                elif key == 'pot' or key == 'pay': self.dm.parse_generic(new_content, key=key, is_path=False)
                
                # [PERSISTENCE] Save to disk
                if key in self.loaded_paths and self.loaded_paths[key]:
                    try:
                        with open(self.loaded_paths[key], 'w', encoding='ISO-8859-1') as f:
                            f.write(new_content)
                    except Exception as e:
                        print(f"Error saving to disk: {e}")
                        
                self.db.save_state(self.dm)
                self.refresh_labels()
                messagebox.showinfo("Fatto", "Modifiche salvate con successo.")
            except Exception as e: messagebox.showerror("Errore Salvataggio", str(e))

        EditorWindow(self, title, content, save_callback, max_rows=limit)

    def refresh_labels(self):
        if 'main' in self.file_boxes: 
            count = 0
            for cls in self.dm.db_main: count += 1 
            txt = f"Caricati: {count} classi" if count > 0 else "Nessun dato"
            self.file_boxes['main']['label'].config(text=txt, fg="#2E7D32" if count else "#888")
            
        if 'pers' in self.file_boxes:
            c = len(self.dm.db_pers)
            self.file_boxes['pers']['label'].config(text=f"Docenti: {c}" if c else "Nessun dato", fg="#2E7D32" if c else "#888")
        
        if 'sos' in self.file_boxes:
            c = len(self.dm.db_sos)
            self.file_boxes['sos']['label'].config(text=f"Docenti: {c}" if c else "Nessun dato", fg="#2E7D32" if c else "#888")
            
        for k in ['alt', 'pot', 'pay']:
            if k in self.file_boxes:
                has_data = False
                if k == 'alt': has_data = any(self.dm.db_alt[d] for d in self.dm.db_alt)
                elif k == 'pot': has_data = any(self.dm.db_pot[d] for d in self.dm.db_pot)
                elif k == 'pay': has_data = any(self.dm.db_pay[d] for d in self.dm.db_pay)
                self.file_boxes[k]['label'].config(text="Dati caricati" if has_data else "Nessun dato", fg="#2E7D32" if has_data else "#888")

    def create_styled_text_box(self, parent, title, r, c, color, hint=None, bg_color="white"):
        f = tk.Frame(parent, bg=bg_color, highlightbackground=color, highlightthickness=1, padx=5, pady=2)
        f.grid(row=r, column=c, sticky="nsew", padx=5, pady=2) # nsew to fill the grid cell
        parent.columnconfigure(c, weight=1)
        
        tk.Label(f, text=title, bg=bg_color, font=("Segoe UI", 10, "bold"), fg="#2c3e50").pack(anchor="w")
        
        # Text area similar to the original but inside the styled frame
        txt = tk.Text(f, height=5, font=("Arial", 14, "bold"), bd=1, relief="solid", highlightthickness=0, bg=bg_color, fg="black", insertbackground="black")
        txt.pack(fill="both", expand=True, padx=2, pady=5)
        
        if hint: 
            tk.Label(f, text=hint, bg=bg_color, fg="#7f8c8d", font=("Segoe UI", 8)).pack(anchor="w", padx=2)
        return txt
    
    def on_closing(self):
        # [FIX 4.7] Pulizia file temporanei HTML generati durante la sessione
        for temp_path in self.temp_files:
            try: os.unlink(temp_path)
            except: pass
        
        # Le caselle giornaliere (docenti assenti, permessi brevi, classi assenti) non vengono
        # memorizzate tra sessioni: ad ogni riavvio l'applicazione si presenta ripulita
        self.dm.text_absent = ""
        self.dm.text_short = ""
        self.dm.text_classes = ""
        try: self.dm.total_teachers = int(self.entry_total_teachers.get())
        except: pass
        self.db.save_state(self.dm)
        self.destroy()

    # --- STATS LOGIC ---
    
    def update_stats_display(self):
        w, m, t, d, det = self.db.get_stats()
        self.lbl_stats.config(text=f"Settimanale: {w}%  |  Mensile: {m}%  |  Totale: {t}%")
    
    def reset_stats(self):
        if messagebox.askyesno("Conferma", "Vuoi davvero azzerare tutte le statistiche?"):
            self.db.reset_stats()
            self.update_stats_display()
            
    def print_stats(self):
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
            with tempfile.NamedTemporaryFile(delete=False, suffix='.html', mode='w', encoding='utf-8') as f: f.write(html)
            self.temp_files.append(f.name)  # [FIX 4.7] Traccia per pulizia
            webbrowser.open('file://' + f.name)
        except Exception as e: messagebox.showerror("Errore", str(e))

    def export_db_file(self):
        path = filedialog.asksaveasfilename(defaultextension=".sqlite", filetypes=[("SQLite DB", "*.sqlite")])
        if path:
            try: self.db.export_db(path); messagebox.showinfo("Export", "Database esportato correttamente!")
            except Exception as e: messagebox.showerror("Export", str(e))
            
    def import_db_file(self):
        path = filedialog.askopenfilename(filetypes=[("SQLite DB", "*.sqlite")])
        if path:
            confirm = messagebox.askyesno("Conferma", "Se importi un database, i dati attuali verranno sovrascritti. Continuare?")
            if confirm:
                try:
                    self.db.import_db(path)
                    if self.db.load_state(self.dm):
                        self.refresh_labels()
                        self.txt_absent.delete("1.0", "end")
                        self.txt_short_leave.delete("1.0", "end")
                        self.txt_classes_absent.delete("1.0", "end")
                        self.entry_total_teachers.delete(0, "end"); self.entry_total_teachers.insert(0, str(self.dm.total_teachers))
                        self.update_stats_display()
                        messagebox.showinfo("Import", "Database importato con successo!")
                    else: messagebox.showerror("Errore", "Impossibile caricare il database importato.")
                except Exception as e: messagebox.showerror("Errore Import", str(e))

    def clear_table(self):
        for item in self.tree.get_children(): self.tree.delete(item)

    def run_calculation(self):
        self.clear_table(); day = self.combo_day.get(); raw_input = self.entry_teacher.get().strip()
        
        # [UI TARGETS] Teachers explicitly requested for substitution (UI Display)
        ui_targets_set = set()
        if raw_input:
             ui_targets_set = {to_slug(x) for x in re.split(r'[,;]+', raw_input) if len(x) > 2}
        
        raw_bl = self.txt_absent.get("1.0", "end"); 
        # [STATS TARGETS] Teachers explicitly listed as absent (Statistics)
        stats_targets_list = [to_slug(x) for x in re.split(r'[\n,]+', raw_bl) if len(x) > 2]
        stats_targets_set = set(stats_targets_list)
        
        # [PROCESSING TARGETS] Union of both (we need to process all of them)
        # If a teacher is absent but no sub requested, we still need to process to count stats.
        # If a teacher is present but sub requested, we process to show UI.
        processing_targets = list(ui_targets_set.union(stats_targets_set))
        
        if not processing_targets: return messagebox.showerror("Errore", "Inserisci almeno un docente (da sostituire o assente)!")
        
        # Blacklist for engine is the list of ACTUALLY absent teachers
        blacklist = list(stats_targets_list)
        
        short_leaves_map = {}; raw_short = self.txt_short_leave.get("1.0", "end")
        for line in raw_short.splitlines():
            for entry in line.strip().split(','):
                parts = entry.strip().split()
                if len(parts) >= 2:
                    s_slug = to_slug(parts[0]); hours = [int(p) for p in parts[1:] if p.isdigit()]
                    if s_slug and hours:
                        if s_slug not in short_leaves_map: short_leaves_map[s_slug] = []
                        short_leaves_map[s_slug].extend(hours)
                        
        # Ensure teachers to be replaced are treated as 'busy' or 'absent' by the engine so they don't sub themselves
        # if they are not already in the blacklist
        for t in processing_targets:
            if t not in short_leaves_map and t not in blacklist: blacklist.append(t)
            
        raw_cls_abs = self.txt_classes_absent.get("1.0", "end"); absent_classes_slugs = [normalize_class(x) for x in re.split(r'[\n,]+', raw_cls_abs) if len(x) > 1]
        
        results = self.engine.find_substitutions(day, processing_targets, blacklist, short_leaves_map, absent_classes_slugs)
        # Sort results by Class (x[0]) and Hour (x[1])
        results.sort(key=lambda x: (x[0], x[1]))
        
        # STATS ACCUMULATION
        try:
            tot = int(self.entry_total_teachers.get())
            self.dm.total_teachers = tot
            
            # [STATS FIX] Statistics should count both fully absent teachers AND short leaves
            stats_valid_set = stats_targets_set.union(short_leaves_map.keys())
            
            new_details = {}
            if results:
                for r in results:
                    # r: (cls, h, txt, res_txt, tag, teacher_slug)
                    if len(r) >= 6:
                        t_slugs = str(r[5]).split(',')
                        for ts in t_slugs:
                            if ts and ts in stats_valid_set:
                                new_details[ts] = new_details.get(ts, 0) + 1
            
            if new_details:
                self.db.accumulate_daily_stats(new_details, tot)
                self.update_stats_display()
        except Exception as e:
            print(f"Stats Error: {e}")

        if not results: messagebox.showinfo("Info", "Nessun risultato trovato.")
        
        # [FIX 1.3] UI FILTER: se nessun filtro UI specificato, mostra tutti i risultati
        show_all = len(ui_targets_set) == 0
        for r in results:
            if len(r) >= 6:
                t_slugs = str(r[5]).split(',')
                if show_all or any(ts in ui_targets_set for ts in t_slugs):
                    self.tree.insert("", "end", values=r[:4], tags=(r[4],))

    def export_excel(self):
        if not HAS_PANDAS: return messagebox.showerror("Errore", "Libreria pandas non trovata.")
        data = []; 
        for item in self.tree.get_children(): data.append(self.tree.item(item)['values'])
        if not data: return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not path: return
        
        # [FIX 2.6] Lettura dati dal thread principale PRIMA di lanciare il worker (thread-safety)
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
                    if s_slug: short_leaves_set.add(s_slug)
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
                    df_main.to_excel(writer, index=False, startrow=len(summary_data)+2)
                self.after(0, lambda: messagebox.showinfo("Fatto", "Excel salvato."))
            except Exception as e: self.after(0, lambda: messagebox.showerror("Errore", str(e)))
        threading.Thread(target=worker, daemon=True).start()

    def ask_export_word_mode(self):
        """Apre un dialog modale per consentire all'utente di scegliere se esportare Word a colori o in B/N."""
        choice = [None]
        dialog = tk.Toplevel(self)
        dialog.title("Esportazione Word")
        dialog.geometry("460x230")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        dialog.configure(bg="#F8F9FA")
        
        # Centra la finestra rispetto alla finestra principale
        dialog.update_idletasks()
        try:
            pw, ph = self.winfo_width(), self.winfo_height()
            px, py = self.winfo_rootx(), self.winfo_rooty()
            dx = px + (pw - 460) // 2
            dy = py + (ph - 230) // 2
            dialog.geometry(f"+{max(0, dx)}+{max(0, dy)}")
        except Exception:
            pass
        
        tk.Label(dialog, text="Esportazione Documento Word", font=("Helvetica", 14, "bold"), bg="#F8F9FA", fg="#2c3e50").pack(pady=(18, 6))
        tk.Label(dialog, text="Come desideri formattare le sostituzioni nel file Word?", font=("Segoe UI", 11), bg="#F8F9FA", fg="#555555").pack(pady=(0, 16))
        
        btn_frame = tk.Frame(dialog, bg="#F8F9FA")
        btn_frame.pack(pady=5)
        
        def choose(val):
            choice[0] = val
            dialog.destroy()
            
        btn_color = tk.Button(btn_frame, text="🎨  A Colori\n(come Anteprima PDF)", font=("Segoe UI", 10, "bold"),
                              bg="#E3F2FD", fg="#1565C0", activebackground="#BBDEFB",
                              relief="groove", padx=16, pady=8, cursor="hand2", command=lambda: choose("color"))
        btn_color.pack(side="left", padx=10)
        
        btn_bw = tk.Button(btn_frame, text="📄  Bianco e Nero\n(Stampa B/N standard)", font=("Segoe UI", 10, "bold"),
                           bg="#ECEFF1", fg="#37474F", activebackground="#CFD8DC",
                           relief="groove", padx=16, pady=8, cursor="hand2", command=lambda: choose("bw"))
        btn_bw.pack(side="left", padx=10)
        
        btn_cancel = tk.Button(dialog, text="Annulla", font=("Segoe UI", 10),
                               bg="#F5F5F5", fg="#7f8c8d", relief="flat", command=lambda: choose(None))
        btn_cancel.pack(pady=(14, 10))
        
        dialog.bind("<Escape>", lambda e: choose(None))
        self.wait_window(dialog)
        return choice[0]

    def export_word(self):
        if not HAS_DOCX: return messagebox.showerror("Errore", "python-docx non trovata.")
        if not self.tree.get_children():
            return messagebox.showinfo("Info", "Nessuna riga di sostituzione presente nella tabella da esportare.")
            
        # Selezione modalità (A Colori o Bianco e Nero)
        mode = self.ask_export_word_mode()
        if not mode:
            return
            
        path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
        if not path: return
        
        # [FIX 2.6] Lettura dati dal thread principale PRIMA di lanciare il worker (thread-safety)
        teacher_txt = self.entry_teacher.get().upper(); day_txt = self.combo_day.get()
        
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
                    if s_slug: short_leaves_set.add(s_slug)
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
                self.after(0, lambda: messagebox.showinfo("Fatto", f"Documento Word ({mode_label}) salvato con successo."))
            except Exception as e:
                err_details = traceback.format_exc()
                print(f"Word Export Error:\n{err_details}")
                self.after(0, lambda: messagebox.showerror("Errore Esportazione Word", f"Si è verificato un errore:\n\n{str(e)}\n\nDettagli:\n{err_details[:300]}..."))
        threading.Thread(target=worker, daemon=True).start()

    def print_preview(self):
        txt_ass = self.txt_absent.get("1.0", "end").strip().replace('\n', '<br>')
        txt_per = self.txt_short_leave.get("1.0", "end").strip().replace('\n', '<br>')
        txt_cla = self.txt_classes_absent.get("1.0", "end").strip().replace('\n', '<br>')
        
        # Calc Off Site
        raw_input = self.entry_teacher.get().strip(); 
        ui_targets_set = {to_slug(x) for x in re.split(r'[,;]+', raw_input) if len(x) > 2} if raw_input else set()
        raw_bl = self.txt_absent.get("1.0", "end"); 
        stats_targets_set = {to_slug(x) for x in re.split(r'[\n,]+', raw_bl) if len(x) > 2}
        raw_short = self.txt_short_leave.get("1.0", "end")
        short_leaves_set = set()
        for line in raw_short.splitlines():
            for entry in line.strip().split(','):
                parts = entry.strip().split()
                if len(parts) >= 2 and any(p.isdigit() for p in parts[1:]):
                    s_slug = to_slug(parts[0])
                    if s_slug: short_leaves_set.add(s_slug)
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
        for k, v in Configuration.COLORS.items(): html += f"tr.{k} td {{ background-color: {v}; }}\n"
            
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
            v = self.tree.item(item)['values']; t = self.tree.item(item)['tags']; tag = t[0] if t else ""
            sub = str(v[3]).replace("\n", "<br>")
            html += f"<tr class='{tag}'><td><b>{v[0]}</b></td><td style='text-align:center'>{v[1]}</td><td>{v[2]}</td><td>{sub}</td></tr>"
            
        html += "</table><div style='text-align:center; margin-top:20px'>© Mario Carta</div><script>window.print()</script></body></html>"
        
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix='.html', mode='w', encoding='utf-8') as f: f.write(html)
            self.temp_files.append(f.name)  # [FIX 4.7] Traccia per pulizia
            webbrowser.open('file://' + f.name)
        except Exception as e: messagebox.showerror("Errore", str(e))

if __name__ == "__main__":
    # Ensure we use a clean environment for Tcl/Tk if possible
    # but don't force bad paths.
    
    try:
        app = App()
        app.mainloop()
    except Exception as e:
        import traceback
        log_path = os.path.expanduser("~/Desktop/gestione_crash.log")
        with open(log_path, "w", encoding="utf-8") as f:
            f.write(f"SI È VERIFICATO UN ERRORE ALL'AVVIO:\n{str(e)}\n\n")
            f.write(traceback.format_exc())
        
        # Try to show a GUI error if possible
        try:
            import tkinter.messagebox as mb
            mb.showerror("Errore Avvio", f"Il programma non è riuscito a partire.\nUn log è stato salvato sul Desktop: gestione_crash.log\n\nErrore: {str(e)}")
        except:
            pass
        sys.exit(1)
