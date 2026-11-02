"""
gestione_sostituzioni.config.settings
Configurazione centralizzata: costanti, colori, tema, percorsi.
"""

import os
import sys
import logging

logger = logging.getLogger("GestioneSostituzioni.Config")


class Configuration:
    """Configurazione centralizzata dell'applicazione."""

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
        'sostegno': '#D1C4E9',
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
        """Restituisce il percorso scrivibile del database per la piattaforma corrente."""
        if sys.platform == 'darwin':
            lib_path = os.path.expanduser('~/Library/Application Support/GestioneSostituzioni')
        elif sys.platform == 'win32':
            lib_path = os.path.join(os.environ.get('APPDATA', os.path.expanduser('~')), 'GestioneSostituzioni')
        else:
            lib_path = os.path.expanduser('~/.config/GestioneSostituzioni')
        if not os.path.exists(lib_path):
            os.makedirs(lib_path, exist_ok=True)
        return os.path.join(lib_path, Configuration.DB_NAME)

    @staticmethod
    def resource_path(relative_path: str) -> str:
        """Restituisce il percorso della risorsa, compatibile con PyInstaller."""
        try:
            base_path = sys._MEIPASS
        except AttributeError:
            base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base_path, relative_path)
