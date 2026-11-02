"""
gestione_sostituzioni.utils
Funzioni di utilità: pulizia stringhe, slug, riconoscimento materie,
normalizzazione classi e parsing ore.
"""

import re
import logging

from .config import Configuration

logger = logging.getLogger("GestioneSostituzioni.Utils")


def clean_str(s: str) -> str:
    """Rimuove spazi e virgolette da una stringa."""
    return s.strip().replace('"', '') if s else ""


def to_slug(str_val: str) -> str:
    """
    Converte un nome docente in uno slug univoco.
    Include le prime 3 lettere del nome per disambiguare omonimie.
    Gestisce cognomi composti con prefissi (De, Di, Lo, La, ecc.).
    """
    if not str_val:
        return ""
    no_parens = re.sub(r'\([^)]*\)', ' ', str_val)
    clean = re.sub(r'[^A-Z0-9\s]', ' ', no_parens.upper())
    parts = clean.split()
    if not parts:
        return ""
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
    """
    Determina se una parola è il nome di una materia scolastica.
    Usa match esatto per evitare falsi positivi su cognomi reali
    (es. Francese, Converso).
    """
    w = word.upper().replace('.', '').replace(',', '')
    EXACT_MATCH = {
        "ITALIANO", "STORIA", "INGLESE", "MATEMATICA", "RELIGIONE",
        "SOSTEGNO", "DIRITTO", "ECONOMIA", "INFORMATICA", "SCIENZE",
        "CHIMICA", "FISICA", "LETTERE", "LABORATORIO", "ENOGASTRONOMIA",
        "CUCINA", "SALA", "VENDITA", "ACCOGLIENZA", "TURISTICA",
        "ALIMENTAZIONE", "FRANCESE", "TEDESCO", "SPAGNOLO", "TIC", "TTP"
    }
    PREFIX_MATCH = {"EDFISICA", "LABOR", "CONV", "GEO"}
    return w in EXACT_MATCH or any(w.startswith(p) for p in PREFIX_MATCH)


def normalize_class(name: str) -> str:
    """
    Normalizza il nome di una classe scolastica.
    Le sigle esplicite di biennio (AB, BB, CB, DB, EB, FB) vengono espanse,
    supportando anche le classi sdoppiate di laboratorio vendita/sala.
    """
    if not name:
        return ""
    n = name.upper().strip()
    n = re.sub(r'\s+[csCS]$', '', n)
    is_vendita = ("VENDITA" in n) or ("SALA" in n)
    # Rimuovi parole descrittive di laboratorio/vendita per isolare la classe base
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



def resolve_school_class(raw_text: str) -> str:
    """
    Risolve ed estrae il nome normalizzato della classe da una cella dell'orario
    (es. orario docenti di sostegno contenente 'SCIENZE 1B', '4 A acc - Accoglienza',
    '3 C SALA DIR E TECN', '1F alimentazione', ecc.).
    Restituisce la classe normalizzata o una stringa vuota se non didattica.
    """
    if not raw_text:
        return ""
    text = ' '.join(raw_text.replace('\n', ' ').strip().split()).upper()
    if any(k in text for k in ['RICEVIMENTO', 'DISPOSIZIONE', 'AULA BES', 'F.S.', 'COORDINAMENTO']):
        return ""

    # Triennio (classi 3, 4, 5)
    m_tri = re.search(r'([3-5])\s*([A-F])(?:\s*(ACC|ENO|SAL|SALA|CUCINA|S))?', text)
    if m_tri:
        year = m_tri.group(1)
        sec = m_tri.group(2)
        spec = m_tri.group(3) or ''

        if any(w in text for w in ['ACC', 'ACCOGLIENZA']):
            return f"{year}ACC" if sec == 'A' else f"{year}{sec}ACC"
        if any(w in text for w in ['SAL', 'SALA']):
            return f"{year}{sec}SAL"
        if any(w in text for w in ['ENO', 'CUCINA', 'ENOGASTRONOMIA']):
            return f"{year}{sec}ENO"
        if spec == 'S':
            return f"{year}{sec}SAL"

        if sec == 'A' and 'ACC' in text:
            return f"{year}ACC"
        if sec == 'A' and ('SAL' in text or spec == 'S'):
            return f"{year}ASAL"
        if sec == 'A':
            return f"{year}AENO"
        if sec == 'B':
            return f"{year}BENO"
        if sec == 'C':
            return f"{year}CENO"
        return f"{year}{sec}"

    # Biennio (classi 1 e 2)
    m_bi = re.search(r'([1-2])\s*([A-F])', text)
    if m_bi:
        year = m_bi.group(1)
        sec = m_bi.group(2)
        if 'SALA' in text and 'CUCINA' not in text:
            return f"{year}{sec}SALA"
        return f"{year}{sec}ABIENNIO"

    return normalize_class(text)


def get_hour(text: str) -> int:
    """
    Estrae il numero dell'ora scolastica da una stringa.
    Supporta sia ore scolastiche ordinali (1-8) che fasce orarie reali (08:00-15:00).
    """
    if not text:
        return 0
    s = text.strip()
    # Tentativo 1: Numero semplice 1-8 (ora scolastica diretta)
    if s.isdigit():
        h = int(s)
        if 1 <= h <= 8:
            return h
    # Tentativo 2: Formato orario HH:MM o HH.MM
    match = re.search(r'(^|\s|:)0?(\d{1,2})[:\.]', s)
    if not match:
        match = re.search(r'^0?(\d{1,2})$', s)
    if match:
        try:
            val = match.group(2) if len(match.groups()) > 1 and match.group(2) else match.group(1)
            h = int(val)
            if 8 <= h <= 15:
                return h - 7
        except Exception:
            logger.debug(f"Impossibile interpretare ora da: '{text}'")
            return 0
    return 0
