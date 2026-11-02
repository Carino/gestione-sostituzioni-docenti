#!/usr/bin/env python3
"""
Gestione Sostituzioni v3.3 — Entry Point Modulare
Avvia l'applicazione dalla struttura modulare gestione_sostituzioni/.
"""

import os
import sys
import logging

# Assicura che il path del progetto sia nel PYTHONPATH
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from gestione_sostituzioni.logging_config import setup_logging
from gestione_sostituzioni.app import App

logger = logging.getLogger("GestioneSostituzioni")


def main():
    """Entry point principale dell'applicazione."""
    # Inizializza il logging strutturato
    setup_logging(level=logging.INFO)
    logger.info("=== Gestione Sostituzioni v3.3 — Avvio ===")

    try:
        app = App()
        app.mainloop()
    except Exception as e:
        logger.critical(f"Errore fatale all'avvio: {e}", exc_info=True)

        # Salva crash log sul Desktop come fallback
        log_path = os.path.expanduser("~/Desktop/gestione_crash.log")
        try:
            import traceback
            with open(log_path, "w", encoding="utf-8") as f:
                f.write(f"SI È VERIFICATO UN ERRORE ALL'AVVIO:\n{str(e)}\n\n")
                f.write(traceback.format_exc())
        except Exception:
            pass

        # Try to show a GUI error if possible
        try:
            import tkinter.messagebox as mb
            mb.showerror("Errore Avvio",
                         f"Il programma non è riuscito a partire.\n"
                         f"Un log è stato salvato sul Desktop: gestione_crash.log\n\n"
                         f"Errore: {str(e)}")
        except Exception:
            pass
        sys.exit(1)


if __name__ == "__main__":
    main()
