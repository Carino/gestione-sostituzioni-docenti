"""
gestione_sostituzioni.exporters
Helper per l'esportazione documenti (Word, Excel, PDF/HTML).
"""

import logging

logger = logging.getLogger("GestioneSostituzioni.Exporters")

# --- Librerie opzionali ---

try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    import lxml  # noqa: F401 — richiesto da python-docx
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False
    logger.info("python-docx non disponibile — esportazione Word disabilitata")

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False
    logger.info("pandas non disponibile — esportazione Excel disabilitata")


def set_docx_cell_shading(cell, color_hex):
    """Applica un colore di sfondo a una cella di tabella Word (.docx)."""
    if not color_hex:
        return
    fill = color_hex.lstrip('#').upper()
    try:
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill}"/>')
        cell._tc.get_or_add_tcPr().append(shading)
    except Exception as e:
        logger.debug(f"Impossibile applicare sfondo cella Word: {e}")
