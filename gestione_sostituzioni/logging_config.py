"""
gestione_sostituzioni.logging_config
Configurazione centralizzata del logging strutturato.
"""

import os
import sys
import logging
from logging.handlers import RotatingFileHandler


def setup_logging(level=logging.INFO):
    """
    Configura il logging strutturato per l'intera applicazione.
    - Console: messaggi INFO e superiori
    - File: tutti i messaggi (DEBUG incluso), con rotazione a 5 MB
    """
    root_logger = logging.getLogger("GestioneSostituzioni")
    root_logger.setLevel(logging.DEBUG)  # Cattura tutto, i filtri sono sugli handler

    # Formattatore dettagliato
    detailed_fmt = logging.Formatter(
        '%(asctime)s [%(levelname)-8s] %(name)s: %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

    # Formattatore compatto per console
    console_fmt = logging.Formatter(
        '[%(levelname)s] %(message)s'
    )

    # --- Handler Console ---
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(level)
    console_handler.setFormatter(console_fmt)
    root_logger.addHandler(console_handler)

    # --- Handler File (con rotazione) ---
    try:
        if sys.platform == 'darwin':
            log_dir = os.path.expanduser('~/Library/Logs/GestioneSostituzioni')
        elif sys.platform == 'win32':
            log_dir = os.path.join(os.environ.get('APPDATA', os.path.expanduser('~')), 'GestioneSostituzioni', 'logs')
        else:
            log_dir = os.path.expanduser('~/.local/share/GestioneSostituzioni/logs')

        os.makedirs(log_dir, exist_ok=True)
        log_file = os.path.join(log_dir, 'gestione_sostituzioni.log')

        file_handler = RotatingFileHandler(
            log_file, maxBytes=5 * 1024 * 1024, backupCount=3, encoding='utf-8'
        )
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(detailed_fmt)
        root_logger.addHandler(file_handler)

        root_logger.info(f"Logging inizializzato — file di log: {log_file}")
    except Exception as e:
        root_logger.warning(f"Impossibile creare il file di log: {e}. Solo console attiva.")

    return root_logger
