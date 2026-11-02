"""
gestione_sostituzioni.data_manager
Classe DataManager: gestione e parsing dei file CSV scolastici.
"""

import re
import csv
import logging
from typing import List, Tuple

from .config import Configuration
from .utils import clean_str, to_slug, is_subject, normalize_class, get_hour, resolve_school_class

logger = logging.getLogger("GestioneSostituzioni.DataManager")


class DataManager:
    """Gestisce il caricamento, parsing e memorizzazione dei dati scolastici."""

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

    def clear_category(self, key: str):
        """Azzera i dati in memoria per la categoria specificata."""
        if key == 'main':
            self.db_main = {}
        elif key == 'pers':
            self.db_pers = {}
        elif key == 'sos':
            self.db_sos = {}
        elif key == 'alt':
            self.db_alt = {d: {} for d in Configuration.DAYS if d}
        elif key == 'pot':
            self.db_pot = {}
        elif key == 'pay':
            self.db_pay = {}

        if key in self.preview_data:
            del self.preview_data[key]
        logger.info(f"Dati in memoria azzerati per la categoria [{key}]")

    def _read_excel(self, file_path: str) -> Tuple[List[str], str]:
        """
        Legge file Excel (.xlsx, .xlsm, .xls) da Excel 2000 in poi
        e li converte in formato righe/testo CSV con delimitatore ';'.
        """
        import io
        import datetime
        ext = file_path.lower().split('.')[-1]
        rows = []

        if ext in ('xlsx', 'xlsm', 'xltx', 'xltm'):
            import openpyxl
            wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
            sheet = wb.active
            for r in sheet.iter_rows(values_only=True):
                formatted_row = []
                for cell in r:
                    if cell is None:
                        formatted_row.append('')
                    elif isinstance(cell, float) and cell.is_integer():
                        formatted_row.append(str(int(cell)))
                    elif isinstance(cell, (datetime.datetime, datetime.time)):
                        formatted_row.append(cell.strftime('%H:%M'))
                    else:
                        formatted_row.append(str(cell).strip())
                if any(formatted_row):
                    rows.append(formatted_row)
            wb.close()
        elif ext in ('xls', 'xlt'):
            import xlrd
            wb = xlrd.open_workbook(file_path)
            sheet = wb.sheet_by_index(0)
            for r_idx in range(sheet.nrows):
                formatted_row = []
                for c_idx in range(sheet.ncols):
                    c_type = sheet.cell_type(r_idx, c_idx)
                    c_val = sheet.cell_value(r_idx, c_idx)
                    if c_type == xlrd.XL_CELL_EMPTY or c_val is None:
                        formatted_row.append('')
                    elif c_type == xlrd.XL_CELL_NUMBER:
                        if float(c_val).is_integer():
                            formatted_row.append(str(int(c_val)))
                        else:
                            formatted_row.append(str(c_val))
                    elif c_type == xlrd.XL_CELL_DATE:
                        try:
                            dt = xlrd.xldate_as_datetime(c_val, wb.datemode)
                            formatted_row.append(dt.strftime('%H:%M'))
                        except Exception:
                            formatted_row.append(str(c_val))
                    else:
                        formatted_row.append(str(c_val).strip())
                if any(formatted_row):
                    rows.append(formatted_row)

        out = io.StringIO()
        writer = csv.writer(out, delimiter=';')
        for row in rows:
            writer.writerow(row)
        raw = out.getvalue()
        logger.info(f"File Excel convertito con successo ({len(rows)} righe): {file_path}")
        return raw.splitlines(), raw

    def _get_content(self, source: str, is_path: bool) -> Tuple[List[str], str]:
        """
        Ottiene il contenuto da un file (CSV o Excel) o da una stringa.
        Per i file Excel (.xlsx, .xls), converte automaticamente le celle in CSV.
        Per i file CSV, prova diversi encoding in cascata.
        """
        if is_path:
            lower = source.lower()
            if lower.endswith(('.xlsx', '.xlsm', '.xltx', '.xltm', '.xls', '.xlt')):
                try:
                    return self._read_excel(source)
                except Exception as e:
                    logger.error(f"Errore lettura file Excel '{source}': {e}", exc_info=True)
                    return [], ""

            encodings = ['utf-8-sig', 'utf-8', 'cp1252', 'ISO-8859-1']
            for enc in encodings:
                try:
                    with open(source, 'r', encoding=enc) as f:
                        content = f.read()
                        logger.debug(f"File letto con encoding '{enc}': {source}")
                        return content.splitlines(), content
                except (UnicodeDecodeError, UnicodeError):
                    continue
                except Exception as e:
                    logger.error(f"Errore lettura file '{source}': {e}")
                    return [], ""
            logger.warning(f"Nessun encoding valido per il file: {source}")
            return [], ""
        else:
            return source.splitlines(), source

    def read_csv_rows(self, lines: List[str]) -> List[List[str]]:
        """Parsing delle righe CSV con rilevamento automatico del delimitatore."""
        if not lines:
            return []
        try:
            # Try to detect dialect
            sample = "\n".join(lines[:5])
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=';,')
            except csv.Error:
                dialect = None

            if dialect:
                reader = csv.reader(lines, dialect)
            else:
                # Fallback default
                sep = ';' if lines[0].count(';') >= lines[0].count(',') else ','
                reader = csv.reader(lines, delimiter=sep)

            return list(reader)
        except Exception as e:
            logger.error(f"Errore parsing CSV: {e}")
            # Ultra-fallback
            sep = ';' if lines[0].count(';') >= lines[0].count(',') else ','
            return [line.split(sep) for line in lines]

    def parse_main(self, source: str, is_path: bool = True) -> int:
        """Parsing dell'orario classi. Restituisce il numero di classi caricate."""
        self.db_main = {}
        lines, raw = self._get_content(source, is_path)
        self.preview_data['main'] = raw
        if not lines:
            return 0

        rows = self.read_csv_rows(lines)

        curr_class = None
        last_h = 0
        day_map = {}
        count = 0
        for row in rows:
            line_str = " ".join(row)
            upper = line_str.upper()
            cols = [clean_str(c) for c in row]

            if "LUNED" in upper or "MARTED" in upper:
                for i, c in enumerate(cols):
                    for d in Configuration.DAYS:
                        if d and d[:3] in c.upper():
                            day_map[i] = d
                            break
            if "CLASSE" in upper:
                is_vendita_or_sala = ("VENDITA" in upper) or ("SALA" in upper)
                m = re.search(r'CLASSE\s+([^,;\(]+)', line_str, re.IGNORECASE)
                if m:
                    base_class = normalize_class(m.group(1))
                    is_biennio = len(base_class) >= 2 and base_class[0] in "12"
                    # Riconosci come classe sdoppiata di laboratorio vendita/sala:
                    # 1. Se l'intestazione contiene VENDITA o SALA
                    # 2. Se la classe di biennio è già presente nell'orario (secondo blocco sdoppiato)
                    if is_biennio and (is_vendita_or_sala or base_class in self.db_main):
                        curr_class = f"{base_class} VENDITA" if not base_class.endswith(" VENDITA") else base_class
                    else:
                        curr_class = base_class
                    if curr_class not in self.db_main:
                        self.db_main[curr_class] = {}
                        count += 1
                    last_h = 0

            if curr_class and len(cols) > 1:
                h = get_hour(cols[0])
                use_def = len(day_map) == 0
                if h > 0:
                    last_h = h
                    for i in range(1, len(cols)):
                        if i >= len(cols):
                            break
                        day = Configuration.DAYS[i] if use_def else day_map.get(i)
                        if day and len(cols[i]) > 1:
                            if day not in self.db_main[curr_class]:
                                self.db_main[curr_class][day] = {}
                            existing = self.db_main[curr_class][day].get(h)
                            if existing and cols[i].upper() not in existing['text'].upper():
                                existing['text'] += f" {cols[i]}"
                            elif not existing:
                                self.db_main[curr_class][day][h] = {'text': cols[i], 'blocked': []}
                elif "COMPR" in upper and last_h > 0:
                    for i in range(1, len(cols)):
                        if i >= len(cols):
                            break
                        day = Configuration.DAYS[i] if use_def else day_map.get(i)
                        if day and len(cols[i]) > 1:
                            if day in self.db_main[curr_class] and last_h in self.db_main[curr_class][day]:
                                self.db_main[curr_class][day][last_h]['text'] += f" (Compr: {cols[i]})"
                                # Compresenti separati dai 'blocked': i compresenti
                                # sono docenti disponibili come sostituti interni, non da bloccare
                                toks = [t for t in re.split(r'[\s\-\/]+', cols[i].upper()) if len(t) > 1 and not is_subject(t)]
                                if 'compresenti' not in self.db_main[curr_class][day][last_h]:
                                    self.db_main[curr_class][day][last_h]['compresenti'] = []
                                self.db_main[curr_class][day][last_h]['compresenti'].extend(toks)

        logger.info(f"Orario classi caricato: {count} classi")
        return count

    def parse_pers(self, source: str, key='pers', is_path: bool = True) -> int:
        """
        Parsing dell'orario personale dei docenti.
        Supporta:
        1. Formato a blocchi orizzontali (es. 'DOCENTE;Ora;Lunedì;Martedì;...'):
           - Prima riga: docente in col 0 e prima ora in col 1
           - Seconda riga: materia (es. ITALIANO) in col 0 e seconda ora in col 1
           - Righe successive: ore successive (col 0 vuota o materia)
        2. Formato a blocchi verticali con etichetta 'DOCENTE [NOME]' e riga 'Ora;Lunedì;...'
        3. Formato legacy a riga singola (DOCENTE; 8:00; 1A; 9:00; 2A...)
        Restituisce il numero di slot caricati.
        """
        if key == 'sos':
            return self.parse_sos(source, is_path=is_path)

        target = {}
        lines, raw = self._get_content(source, is_path)
        self.preview_data[key] = raw
        if not lines:
            return 0

        rows = self.read_csv_rows(lines)
        loaded = 0
        curr_slug = ''
        curr_name = ''
        day_col_map = {}
        non_teaching = {'R', 'DISP', 'DISPOSIZIONE', 'D', 'D.', 'MENSA', 'P', 'POT', 'POTENZIAMENTO', 'RIC', 'RICEVIMENTO'}

        def is_valid_class(val: str) -> bool:
            v = val.strip().upper()
            if not v or v in non_teaching or any(v.startswith(nt) for nt in ['DISP', 'MENSA', 'RICEV']):
                return False
            return any(ch.isdigit() for ch in v) or any(k in v for k in ['BIENNIO', 'SALA', 'ENO', 'ACC'])

        i = 0
        while i < len(rows):
            row = [clean_str(c).strip() for c in rows[i]]
            if not row or not any(row):
                i += 1
                continue

            c0 = row[0]
            c0_upper = c0.upper()

            # 1. Riconoscimento riga intestazione giorni (es. 'DOCENTE;Ora;Lunedì...' o 'Ora;Lunedì...')
            row_days = {}
            for idx, cell in enumerate(row):
                c_clean = clean_str(cell).upper()
                for d in Configuration.DAYS:
                    if d and d[:3] in c_clean:
                        row_days[idx] = d
                        break

            if len(row_days) >= 3:
                day_col_map = row_days
                # Se la riga ha 'DOCENTE' in col 0 (Formato cattedre orizzontale)
                if c0_upper == 'DOCENTE':
                    i += 1
                    if i < len(rows):
                        next_row = [clean_str(c).strip() for c in rows[i]]
                        if next_row and next_row[0]:
                            curr_name = next_row[0]
                            curr_slug = to_slug(curr_name)
                            if curr_slug not in target:
                                target[curr_slug] = {'_name': curr_name}
                            h = get_hour(next_row[1]) if len(next_row) > 1 else 0
                            if h > 0:
                                for idx, d_name in day_col_map.items():
                                    if idx < len(next_row):
                                        val = next_row[idx]
                                        if is_valid_class(val):
                                            cls = normalize_class(val)
                                            if cls:
                                                if d_name not in target[curr_slug]:
                                                    target[curr_slug][d_name] = {}
                                                target[curr_slug][d_name][h] = cls
                                                loaded += 1
                i += 1
                continue

            # 2. Riconoscimento riga docente con prefisso: 'DOCENTE ALIAS' o 'DOCENTE: ALIAS'
            if c0_upper.startswith('DOCENTE ') or c0_upper.startswith('DOCENTE:'):
                teacher_name = c0[7:].strip().lstrip(':').strip()
                if teacher_name:
                    curr_name = teacher_name
                    curr_slug = to_slug(teacher_name)
                    if curr_slug not in target:
                        target[curr_slug] = {'_name': curr_name}
                i += 1
                continue

            # 3. Riga oraria (ora in col 0 o col 1) all'interno di un blocco docente
            h0 = get_hour(c0)
            h1 = get_hour(row[1]) if len(row) > 1 else 0

            if h0 > 0 and curr_slug and day_col_map:
                for idx, d_name in day_col_map.items():
                    if idx < len(row):
                        val = row[idx]
                        if is_valid_class(val):
                            cls = normalize_class(val)
                            if cls:
                                if d_name not in target[curr_slug]:
                                    target[curr_slug][d_name] = {}
                                target[curr_slug][d_name][h0] = cls
                                loaded += 1
                i += 1
                continue

            if h1 > 0 and curr_slug and day_col_map:
                for idx, d_name in day_col_map.items():
                    if idx < len(row):
                        val = row[idx]
                        if is_valid_class(val):
                            cls = normalize_class(val)
                            if cls:
                                if d_name not in target[curr_slug]:
                                    target[curr_slug][d_name] = {}
                                target[curr_slug][d_name][h1] = cls
                                loaded += 1
                i += 1
                continue

            # 4. Formato legacy a riga singola (DOCENTE; 8:00; 1A; 9:00; 2A...)
            ignore = ['DOCENTE', 'MATERIA', 'ORA'] + Configuration.SUBJECT_BLACKLIST
            if len(c0) > 2 and not any(c0_upper.startswith(x) for x in ignore):
                curr_slug = to_slug(c0)
                if curr_slug not in target:
                    target[curr_slug] = {'_name': c0}
                for d in range(6):
                    day = Configuration.DAYS[d + 1]
                    t_idx = 1 + (d * 2)
                    c_idx = t_idx + 1
                    if c_idx < len(row):
                        h = get_hour(row[t_idx])
                        if h > 0 and len(row[c_idx]) > 1:
                            val = row[c_idx]
                            if is_valid_class(val):
                                cls = normalize_class(val)
                                if cls:
                                    if day not in target[curr_slug]:
                                        target[curr_slug][day] = {}
                                    target[curr_slug][day][h] = cls
                                    loaded += 1
            i += 1

        self.db_pers = target
        active_docenti = sum(1 for v in target.values() if isinstance(v, dict) and any(k != '_name' for k in v))
        logger.info(f"Orario {key} caricato: {loaded} slot da {active_docenti} docenti attivi (su {len(target)} totali)")
        return loaded


    def parse_sos(self, source: str, is_path: bool = True) -> int:
        """
        Parsing dell'orario dei docenti di sostegno.
        Supporta la struttura a blocchi docenti (orizzontale o verticale con orari e materie/classi).
        Restituisce il numero di slot caricati.
        """
        lines, raw = self._get_content(source, is_path)
        self.preview_data['sos'] = raw
        if not lines:
            self.db_sos = {}
            return 0

        rows = self.read_csv_rows(lines)
        target = {}
        loaded = 0
        day_col_map = {}
        curr_slug = None
        curr_name = ""

        for row in rows:
            if not any(row):
                continue

            # Rileva riga intestazione giorni (es. NOME DOCENTE, ORA, LUNEDI, MARTEDI...)
            row_days = {}
            for idx, cell in enumerate(row):
                c_clean = clean_str(cell).upper()
                for d in Configuration.DAYS:
                    if d and d[:3] in c_clean:
                        row_days[idx] = d
                        break
            if len(row_days) >= 3:
                day_col_map = row_days
                continue

            c0 = clean_str(row[0])
            c0_upper = c0.upper().strip()
            is_header_label = c0_upper in ["NOME DOCENTE", "DOCENTE", "MATERIA", "ORA"] or c0_upper.startswith("DOCENTE:") or c0_upper.startswith("NOME DOCENTE:")
            # Se la prima colonna contiene il nome del docente
            if len(c0) > 2 and not is_header_label:
                curr_name = c0
                curr_slug = to_slug(c0)
                if curr_slug not in target:
                    target[curr_slug] = {'_name': curr_name}
                continue

            # Righe orarie per il docente corrente
            if curr_slug and len(row) > 1:
                h = get_hour(row[1])
                if h > 0:
                    for c_idx in range(2, len(row)):
                        val = clean_str(row[c_idx])
                        if not val:
                            continue
                        day = day_col_map.get(c_idx)
                        if not day and (c_idx - 1) < len(Configuration.DAYS):
                            day = Configuration.DAYS[c_idx - 1]
                        if not day:
                            continue

                        cls = resolve_school_class(val)
                        if cls:
                            if day not in target[curr_slug]:
                                target[curr_slug][day] = {}
                            target[curr_slug][day][h] = cls
                            loaded += 1

        self.db_sos = target
        num_teachers = sum(1 for v in target.values() if any(k != '_name' for k in v))
        logger.info(f"Orario sostegno caricato: {loaded} slot da {num_teachers} docenti attivi (su {len(target)} totali)")
        return loaded

    def parse_alt(self, source: str, is_path: bool = True) -> int:
        """Parsing della materia alternativa. Restituisce il numero di slot caricati."""
        self.db_alt = {d: {} for d in Configuration.DAYS if d}
        lines, raw = self._get_content(source, is_path)
        self.preview_data['alt'] = raw
        if not lines:
            return 0

        rows = self.read_csv_rows(lines)
        loaded = 0
        for row in rows:
            cols = [clean_str(c) for c in row]
            for d in range(6):
                day = Configuration.DAYS[d + 1]
                t_idx = d * 2
                c_idx = t_idx + 1
                if c_idx < len(cols):
                    h = get_hour(cols[t_idx])
                    val = cols[c_idx]
                    if h > 0 and len(val) > 2:
                        parts = val.strip().split()
                        if len(parts) >= 2:
                            cls = normalize_class(parts[0])
                            teacher = " ".join([p for p in parts[1:] if not is_subject(p)])
                            if h not in self.db_alt[day]:
                                self.db_alt[day][h] = {}
                            self.db_alt[day][h][cls] = teacher
                            loaded += 1

        logger.info(f"Materia alternativa caricata: {loaded} slot")
        return loaded

    def parse_generic(self, source: str, key='pot', is_path: bool = True) -> int:
        """Parsing del potenziamento o pagamento. Restituisce il numero di slot caricati."""
        target_db = {d: {h: [] for h in range(1, 9)} for d in Configuration.DAYS if d}
        lines, raw = self._get_content(source, is_path)
        self.preview_data[key] = raw
        if not lines:
            return 0

        rows = self.read_csv_rows(lines)
        loaded = 0
        for row in rows:
            cols = [clean_str(c) for c in row]
            for d in range(6):
                day = Configuration.DAYS[d + 1]
                t_idx = d * 2
                n_idx = t_idx + 1
                if n_idx < len(cols):
                    h = get_hour(cols[t_idx])
                    raw_val = cols[n_idx]
                    if h > 0 and len(raw_val) > 2:
                        parts = re.split(r'\s+', raw_val)
                        for n in parts:
                            if len(n) > 2 and not is_subject(n):
                                target_db[day][h].append(n)
                                loaded += 1

        if key == 'pot':
            self.db_pot = target_db
        else:
            self.db_pay = target_db

        logger.info(f"Dati {key} caricati: {loaded} slot")
        return loaded
