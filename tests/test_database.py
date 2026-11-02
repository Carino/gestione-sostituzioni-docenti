"""
tests/test_database.py
Test unitari per la classe DatabaseHandler (persistenza SQLite).
"""

import pytest
import sys
import os
import tempfile
import shutil

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from gestione_sostituzioni.database import DatabaseHandler
from gestione_sostituzioni.data_manager import DataManager
from gestione_sostituzioni.config import Configuration


@pytest.fixture
def temp_db(tmp_path):
    """Crea un DatabaseHandler con un database temporaneo."""
    # Override del percorso DB per usare tmp_path
    db_path = str(tmp_path / "test_db.sqlite")
    handler = DatabaseHandler.__new__(DatabaseHandler)
    handler.db_path = db_path
    handler.setup()
    return handler


class TestDatabaseSetup:
    """Test per l'inizializzazione del database."""

    def test_setup_creates_tables(self, temp_db):
        """Il setup deve creare tutte le tabelle necessarie."""
        with temp_db.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT name FROM sqlite_master WHERE type='table'")
            tables = {row[0] for row in c.fetchall()}
        assert "settings" in tables
        assert "data_store" in tables
        assert "preview_store" in tables
        assert "daily_stats" in tables

    def test_setup_idempotent(self, temp_db):
        """Il setup può essere chiamato più volte senza errori."""
        temp_db.setup()  # Second call
        temp_db.setup()  # Third call
        # No exception = success


class TestDailyStats:
    """Test per le statistiche giornaliere."""

    def test_accumulate_replaces_data(self, temp_db):
        """Premendo PROCEDI più volte nello stesso giorno, i dati devono essere sostituiti."""
        # Prima volta
        temp_db.accumulate_daily_stats({"ROSSI": 3, "BIANCHI": 2}, 50)
        w1, m1, t1, d1, det1 = temp_db.get_stats()

        # Seconda volta (stesso giorno) — deve SOSTITUIRE, non accumulare
        temp_db.accumulate_daily_stats({"ROSSI": 4}, 50)
        w2, m2, t2, d2, det2 = temp_db.get_stats()

        # Le ore di BIANCHI non devono persistere dalla prima invocazione
        # Solo ROSSI con 4 ore deve essere presente
        assert det2.get("ROSSI", 0) == 4
        assert "BIANCHI" not in det2

    def test_stats_empty_initially(self, temp_db):
        w, m, t, d, det = temp_db.get_stats()
        assert w == 0
        assert m == 0
        assert t == 0
        assert d == 0
        assert det == {}

    def test_reset_stats(self, temp_db):
        temp_db.accumulate_daily_stats({"ROSSI": 3}, 50)
        temp_db.reset_stats()
        w, m, t, d, det = temp_db.get_stats()
        assert det == {}
        assert d == 0


class TestSaveLoadState:
    """Test per il salvataggio e caricamento dello stato."""

    def test_save_and_load(self, temp_db):
        dm = DataManager()
        dm.total_teachers = 42
        dm.db_main = {"1A": {"LUNEDÌ": {1: {"text": "ROSSI", "blocked": []}}}}

        temp_db.save_state(dm)

        # Crea un nuovo DataManager per verificare il caricamento
        dm2 = DataManager()
        result = temp_db.load_state(dm2)
        assert result is True
        assert dm2.total_teachers == 42

    def test_load_clears_daily_inputs(self, temp_db):
        """Le caselle giornaliere devono essere vuote al caricamento."""
        dm = DataManager()
        dm.text_absent = "ROSSI"
        dm.text_short = "BIANCHI 1 2"
        dm.text_classes = "3A"
        temp_db.save_state(dm)

        dm2 = DataManager()
        temp_db.load_state(dm2)
        # Le caselle giornaliere devono essere vuote all'avvio
        assert dm2.text_absent == ""
        assert dm2.text_short == ""
        assert dm2.text_classes == ""

    def test_clear_category(self, temp_db):
        """Verifica la cancellazione singola di una categoria dal DB."""
        dm = DataManager()
        dm.db_main = {"1A": {"LUNEDÌ": {1: {"text": "ROSSI", "blocked": []}}}}
        dm.preview_data = {"main": "CLASSE 1A;..."}
        temp_db.save_state(dm)

        # Cancella la categoria 'main'
        temp_db.clear_category("main")

        with temp_db.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT json_data FROM data_store WHERE key = 'main'")
            assert c.fetchone()[0] == "{}"
            c.execute("SELECT content FROM preview_store WHERE key = 'main'")
            assert c.fetchone() is None


class TestExportImport:
    """Test per l'esportazione e importazione del database."""

    def test_export_creates_file(self, temp_db, tmp_path):
        export_path = str(tmp_path / "exported.sqlite")
        temp_db.export_db(export_path)
        assert os.path.exists(export_path)

    def test_import_overwrites(self, temp_db, tmp_path):
        # Crea dati nel DB principale
        temp_db.accumulate_daily_stats({"ROSSI": 5}, 50)

        # Esporta
        export_path = str(tmp_path / "backup.sqlite")
        temp_db.export_db(export_path)

        # Resetta
        temp_db.reset_stats()
        _, _, _, _, det = temp_db.get_stats()
        assert det == {}

        # Importa il backup
        temp_db.import_db(export_path)
        _, _, _, _, det = temp_db.get_stats()
        assert "ROSSI" in det
