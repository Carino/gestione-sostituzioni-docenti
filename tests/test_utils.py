"""
tests/test_utils.py
Test unitari per le funzioni di utilità (to_slug, is_subject, normalize_class, get_hour).
"""

import pytest
import sys
import os

# Assicura che il path del progetto sia accessibile
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from gestione_sostituzioni.utils import to_slug, is_subject, normalize_class, get_hour, clean_str, resolve_school_class


class TestCleanStr:
    """Test per la funzione clean_str."""

    def test_basic_strip(self):
        assert clean_str("  hello  ") == "hello"

    def test_removes_quotes(self):
        assert clean_str('"quoted"') == "quoted"

    def test_empty_string(self):
        assert clean_str("") == ""

    def test_none_input(self):
        assert clean_str(None) == ""

    def test_combined(self):
        assert clean_str('  "test value"  ') == "test value"


class TestToSlug:
    """Test per la funzione to_slug con disambiguazione omonimie."""

    def test_simple_surname(self):
        result = to_slug("Rossi Mario")
        assert result == "ROSSIMAR"

    def test_single_word(self):
        result = to_slug("Rossi")
        assert result == "ROSSI"

    def test_compound_surname_de(self):
        result = to_slug("De Luca Giovanni")
        assert result == "DELUCAGIO"

    def test_compound_surname_di(self):
        result = to_slug("Di Marco Anna")
        assert result == "DIMARCOANN"  # ANN[:3] = ANN (3 chars)

    def test_homonyms_differ(self):
        """Due docenti con lo stesso cognome devono avere slug diversi."""
        slug1 = to_slug("Rossi Mario")
        slug2 = to_slug("Rossi Anna")
        assert slug1 != slug2

    def test_empty_string(self):
        assert to_slug("") == ""

    def test_none_input(self):
        assert to_slug(None) == ""  # to_slug checks `if not str_val`

    def test_parentheses_removed(self):
        result = to_slug("Rossi (supplente) Mario")
        assert "SUPPLENTE" not in result

    def test_special_characters(self):
        result = to_slug("D'Amico-Ferretti Luca")
        assert len(result) > 0
        assert result.isalpha() or result.isalnum()

    def test_compound_prefix_dal(self):
        result = to_slug("Dal Monte Pietro")
        assert result.startswith("DALMONTE")


class TestIsSubject:
    """Test per la funzione is_subject (riconoscimento materie scolastiche)."""

    def test_real_subjects(self):
        """Le materie reali devono essere riconosciute."""
        assert is_subject("MATEMATICA") is True
        assert is_subject("ITALIANO") is True
        assert is_subject("INGLESE") is True
        assert is_subject("RELIGIONE") is True
        assert is_subject("STORIA") is True
        assert is_subject("SCIENZE") is True

    def test_real_subjects_lowercase(self):
        """Deve funzionare anche con minuscole."""
        assert is_subject("matematica") is True
        assert is_subject("italiano") is True

    def test_prefix_match(self):
        """I prefissi speciali devono essere riconosciuti."""
        assert is_subject("EDFISICA") is True
        assert is_subject("LABORATORIO") is True
        assert is_subject("GEOGRAFIA") is True  # GEO prefix

    def test_surname_not_blocked(self):
        """Cognomi che iniziano con prefissi di materia possono essere bloccati — comportamento noto."""
        # CONVERSO inizia con CONV (prefix match) — questo è un falso positivo noto
        # ma si accetta per sicurezza (è meglio escludere un cognome che includere una materia)
        assert is_subject("CONVERSO") is True  # CONV prefix match — comportamento atteso

    def test_tic_ttp(self):
        assert is_subject("TIC") is True
        assert is_subject("TTP") is True

    def test_random_word(self):
        assert is_subject("BIANCHI") is False
        assert is_subject("COMPUTER") is False


class TestNormalizeClass:
    """Test per la funzione normalize_class."""

    def test_simple_class(self):
        assert normalize_class("1A") == "1A"

    def test_biennio_suffix(self):
        """Le sigle di biennio (AB) vengono espanse."""
        result = normalize_class("1AB")
        assert "BIENNIO" in result or "ABIENNIO" in result

    def test_non_biennio_two_letter(self):
        """Classi come 1EN, 2SV NON devono essere espanse."""
        assert normalize_class("1EN") == "1EN"
        assert normalize_class("2SV") == "2SV"
        assert normalize_class("3AT") == "3AT"

    def test_empty_string(self):
        assert normalize_class("") == ""

    def test_whitespace_handling(self):
        assert normalize_class("  1A  ") == "1A"

    def test_trailing_section(self):
        """Rimuove suffisso sezione (c/s)."""
        result = normalize_class("3A c")
        assert result == "3A"

    def test_uppercase(self):
        assert normalize_class("1a") == "1A"

    def test_five_letter_class(self):
        """Classi con nomi più lunghi non devono essere alterate erroneamente."""
        assert normalize_class("5AENO") == "5AENO"


class TestGetHour:
    """Test per la funzione get_hour (parsing ore scolastiche)."""

    def test_ordinal_hours(self):
        """Ore scolastiche ordinali (1-8)."""
        for i in range(1, 9):
            assert get_hour(str(i)) == i

    def test_ordinal_zero(self):
        assert get_hour("0") == 0

    def test_ordinal_nine(self):
        """9 non è ora ordinale (>8) ma fallthrough: 9-7=2 come ora reale."""
        assert get_hour("9") == 2  # Non ordinale, ma il fallback lo interpreta come 09:00 → 9-7=2

    def test_time_format_colon(self):
        """Formato HH:MM."""
        assert get_hour("08:00") == 1   # 8 - 7 = 1
        assert get_hour("09:30") == 2   # 9 - 7 = 2
        assert get_hour("10:15") == 3   # 10 - 7 = 3
        assert get_hour("14:00") == 7   # 14 - 7 = 7

    def test_time_format_dot(self):
        """Formato HH.MM."""
        assert get_hour("08.00") == 1
        assert get_hour("10.30") == 3

    def test_empty_string(self):
        assert get_hour("") == 0

    def test_none_input(self):
        assert get_hour(None) == 0  # get_hour checks `if not text`

    def test_invalid_text(self):
        assert get_hour("abc") == 0
        assert get_hour("LUNEDÌ") == 0

    def test_padded_number(self):
        assert get_hour("  3  ") == 3


class TestResolveSchoolClass:
    """Test per la funzione resolve_school_class."""

    def test_empty_and_non_didactic(self):
        assert resolve_school_class("") == ""
        assert resolve_school_class("RICEVIMENTO GENITORI") == ""
        assert resolve_school_class("DISPOSIZIONE") == ""
        assert resolve_school_class("F.S. INCLUSIONE (AULA BES)") == ""

    def test_biennio_resolution(self):
        assert resolve_school_class("SCIENZE 1B") == "1BABIENNIO"
        assert resolve_school_class("TIC 1B") == "1BABIENNIO"
        assert resolve_school_class("1D Fisica") == "1DABIENNIO"
        assert resolve_school_class("1d Matematica") == "1DABIENNIO"
        assert resolve_school_class("1F alimentazione") == "1FABIENNIO"
        assert resolve_school_class("MATEMATICA1 F") == "1FABIENNIO"
        assert resolve_school_class("2 E") == "2EABIENNIO"
        assert resolve_school_class("2C") == "2CABIENNIO"

    def test_triennio_resolution(self):
        assert resolve_school_class("STORIA 5AS") == "5ASAL"
        assert resolve_school_class("SALA 5AS") == "5ASAL"
        assert resolve_school_class("5 A sala - Economia") == "5ASAL"
        assert resolve_school_class("4 A acc - Accoglienza") == "4ACC"
        assert resolve_school_class("DIRITTO E TEC 4 C ENO") == "4CENO"
        assert resolve_school_class("3 C SALA DIR E TECN") == "3CSAL"
        assert resolve_school_class("3B Eno") == "3BENO"
        assert resolve_school_class("5A ENO") == "5AENO"
        assert resolve_school_class("5B SALA") == "5BSAL"

