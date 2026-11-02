"""
gestione_sostituzioni.database
Classe DatabaseHandler: persistenza dati tramite SQLite.
"""

import os
import json
import shutil
import sqlite3
import logging
from datetime import datetime, timedelta

from .config import Configuration
from .data_manager import DataManager

logger = logging.getLogger("GestioneSostituzioni.Database")


class DatabaseHandler:
    """Gestisce la persistenza dei dati tramite database SQLite."""

    def __init__(self):
        self.db_path = Configuration.get_db_path()

        # [MIGRATION] If old local DB exists and new one doesn't, copy it
        bundled_db = Configuration.resource_path(Configuration.DB_NAME)
        if os.path.exists(bundled_db) and not os.path.exists(self.db_path):
            try:
                shutil.copy2(bundled_db, self.db_path)
                logger.info(f"Database migrato da '{bundled_db}' a '{self.db_path}'")
            except Exception as e:
                logger.warning(f"Impossibile migrare il database: {e}")

    def get_connection(self):
        """Restituisce una connessione al database SQLite."""
        return sqlite3.connect(self.db_path)

    def setup(self):
        """Crea le tabelle necessarie se non esistono."""
        with self.get_connection() as conn:
            c = conn.cursor()
            c.execute('''CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT)''')
            c.execute('''CREATE TABLE IF NOT EXISTS data_store (key TEXT PRIMARY KEY, json_data TEXT)''')
            c.execute('''CREATE TABLE IF NOT EXISTS preview_store (key TEXT PRIMARY KEY, content TEXT)''')
            # New Stats table
            c.execute('''CREATE TABLE IF NOT EXISTS daily_stats (date TEXT PRIMARY KEY, absent_count INTEGER, total_teachers INTEGER, details TEXT)''')
            # Attempt migration for existing DBs
            try:
                c.execute("ALTER TABLE daily_stats ADD COLUMN details TEXT")
            except Exception:
                pass
            conn.commit()
        logger.info("Database inizializzato correttamente")

    def accumulate_daily_stats(self, new_details_dict, tot_teachers):
        """
        Salva le statistiche giornaliere con strategia sostitutiva
        (ogni pressione di PROCEDI nello stesso giorno sostituisce i dati).
        """
        today = datetime.now().strftime("%Y-%m-%d")
        count = len(new_details_dict)
        details_json = json.dumps(new_details_dict)

        with self.get_connection() as conn:
            c = conn.cursor()
            c.execute(
                "INSERT OR REPLACE INTO daily_stats (date, absent_count, total_teachers, details) VALUES (?, ?, ?, ?)",
                (today, count, tot_teachers, details_json)
            )
            conn.commit()
        logger.info(f"Statistiche giornaliere aggiornate: {count} docenti assenti su {tot_teachers} totali")

    def get_stats(self):
        """
        Calcola le statistiche aggregate.
        Returns: (week_avg, month_avg, total_avg, total_hours, teacher_details_map)
        """
        with self.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT date, absent_count, total_teachers, details FROM daily_stats")
            rows = c.fetchall()

        if not rows:
            return (0, 0, 0, 0, {})

        now = datetime.now()
        start_week = (now - timedelta(days=now.weekday())).strftime("%Y-%m-%d")
        start_month = now.strftime("%Y-%m-01")

        w_n, w_d = 0, 0
        m_n, m_d = 0, 0
        t_n, t_d = 0, 0

        agg_details = {}  # "Teacher": hours

        for r in rows:
            dt, count, total, det_str = r
            if total > 0:
                t_n += count
                t_d += total
                if dt >= start_week:
                    w_n += count
                    w_d += total
                if dt >= start_month:
                    m_n += count
                    m_d += total

            # Aggregate hours
            if det_str:
                try:
                    d_map = json.loads(det_str)
                    for t_name, h_count in d_map.items():
                        agg_details[t_name] = agg_details.get(t_name, 0) + h_count
                except Exception as e:
                    logger.debug(f"Errore parsing dettagli statistiche: {e}")

        def safe_pct(n, d):
            return round((n / d) * 100, 2) if d > 0 else 0

        return (safe_pct(w_n, w_d), safe_pct(m_n, m_d), safe_pct(t_n, t_d), t_d, agg_details)

    def reset_stats(self):
        """Azzera tutte le statistiche."""
        with self.get_connection() as conn:
            conn.execute("DELETE FROM daily_stats")
            conn.commit()
        logger.info("Statistiche azzerate")

    def clear_category(self, key: str):
        """Cancella i dati di una singola categoria dal database."""
        empty_map = {
            'main': '{}',
            'pers': '{}',
            'sos': '{}',
            'alt': json.dumps({d: {} for d in Configuration.DAYS if d}),
            'pot': '{}',
            'pay': '{}'
        }
        empty_val = empty_map.get(key, '{}')
        with self.get_connection() as conn:
            conn.execute("INSERT OR REPLACE INTO data_store (key, json_data) VALUES (?, ?)", (key, empty_val))
            conn.execute("DELETE FROM preview_store WHERE key = ?", (key,))
            conn.commit()
        logger.info(f"Dati della categoria '{key}' rimossi dal database")

    def save_state(self, dm: DataManager):
        """Salva lo stato corrente nel database."""
        with self.get_connection() as conn:
            c = conn.cursor()
            c.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", ('absent', ''))
            c.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", ('short', ''))
            c.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", ('classes', ''))
            c.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)",
                       ('total_teachers', str(dm.total_teachers)))

            c.execute("INSERT OR REPLACE INTO data_store (key, json_data) VALUES (?, ?)",
                       ('main', json.dumps(dm.db_main)))
            c.execute("INSERT OR REPLACE INTO data_store (key, json_data) VALUES (?, ?)",
                       ('pers', json.dumps(dm.db_pers)))
            c.execute("INSERT OR REPLACE INTO data_store (key, json_data) VALUES (?, ?)",
                       ('sos', json.dumps(dm.db_sos)))
            c.execute("INSERT OR REPLACE INTO data_store (key, json_data) VALUES (?, ?)",
                       ('alt', json.dumps(dm.db_alt)))
            c.execute("INSERT OR REPLACE INTO data_store (key, json_data) VALUES (?, ?)",
                       ('pot', json.dumps(dm.db_pot)))
            c.execute("INSERT OR REPLACE INTO data_store (key, json_data) VALUES (?, ?)",
                       ('pay', json.dumps(dm.db_pay)))

            c.execute("DELETE FROM preview_store")
            for k, v in dm.preview_data.items():
                c.execute("INSERT OR REPLACE INTO preview_store (key, content) VALUES (?, ?)", (k, v))
            conn.commit()
        logger.debug("Stato applicazione salvato nel database")

    def load_state(self, dm: DataManager) -> bool:
        """
        Carica lo stato dal database.
        Le caselle giornaliere sono sempre pulite all'avvio.
        """
        if not os.path.exists(self.db_path):
            return False
        try:
            with self.get_connection() as conn:
                c = conn.cursor()
                curr = c.execute("SELECT key, value FROM settings")
                settings = dict(curr.fetchall())
                # Caselle giornaliere sempre pulite all'avvio
                dm.text_absent = ""
                dm.text_short = ""
                dm.text_classes = ""
                try:
                    dm.total_teachers = int(settings.get('total_teachers', 0))
                except Exception:
                    dm.total_teachers = 0

                curr = c.execute("SELECT key, content FROM preview_store")
                dm.preview_data = dict(curr.fetchall())

                # Re-parse from raw content to ensure robustness
                if 'main' in dm.preview_data:
                    dm.parse_main(dm.preview_data['main'], is_path=False)
                if 'pers' in dm.preview_data:
                    dm.parse_pers(dm.preview_data['pers'], key='pers', is_path=False)
                if 'sos' in dm.preview_data:
                    dm.parse_sos(dm.preview_data['sos'], is_path=False)
                if 'alt' in dm.preview_data:
                    dm.parse_alt(dm.preview_data['alt'], is_path=False)
                if 'pot' in dm.preview_data:
                    dm.parse_generic(dm.preview_data['pot'], key='pot', is_path=False)
                if 'pay' in dm.preview_data:
                    dm.parse_generic(dm.preview_data['pay'], key='pay', is_path=False)

            logger.info("Stato applicazione caricato dal database")
            return True
        except Exception as e:
            logger.error(f"Errore caricamento stato dal database: {e}", exc_info=True)
            return False

    def export_db(self, destination_path: str):
        """Esporta il database su un file esterno."""
        shutil.copy2(self.db_path, destination_path)
        logger.info(f"Database esportato in: {destination_path}")

    def import_db(self, source_path: str):
        """Importa un database da un file esterno."""
        shutil.copy2(source_path, self.db_path)
        logger.info(f"Database importato da: {source_path}")
