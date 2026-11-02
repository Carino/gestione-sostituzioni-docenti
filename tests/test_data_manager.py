"""
tests/test_data_manager.py
Test unitari per la classe DataManager (parsing CSV).
"""

import pytest
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from gestione_sostituzioni.data_manager import DataManager


class TestDataManagerInit:
    """Test per l'inizializzazione del DataManager."""

    def test_default_state(self):
        dm = DataManager()
        assert dm.db_main == {}
        assert dm.db_pers == {}
        assert dm.db_sos == {}
        assert dm.text_absent == ""
        assert dm.text_short == ""
        assert dm.text_classes == ""
        assert dm.total_teachers == 0

    def test_db_alt_initialized(self):
        dm = DataManager()
        # db_alt deve avere una chiave per ogni giorno
        assert "LUNEDÌ" in dm.db_alt
        assert "VENERDÌ" in dm.db_alt
        assert None not in dm.db_alt

    def test_clear_category_all_keys(self):
        dm = DataManager()
        dm.db_main = {"1A": {}}
        dm.db_pers = {"ROSSI": {}}
        dm.db_sos = {"ATZORI": {}}
        dm.db_pot = {"LUNEDÌ": {1: ["GIALLI"]}}
        dm.db_pay = {"LUNEDÌ": {1: ["NERI"]}}
        dm.preview_data = {"main": "data", "pers": "data", "sos": "data", "pot": "data", "pay": "data"}

        dm.clear_category("main")
        assert dm.db_main == {}
        assert "main" not in dm.preview_data
        assert "pers" in dm.preview_data

        dm.clear_category("sos")
        assert dm.db_sos == {}
        assert "sos" not in dm.preview_data

        dm.clear_category("pot")
        assert dm.db_pot == {}
        assert "pot" not in dm.preview_data


class TestReadCsvRows:
    """Test per il parsing CSV con rilevamento automatico del delimitatore."""

    def test_semicolon_delimiter(self):
        dm = DataManager()
        lines = ["A;B;C", "1;2;3"]
        rows = dm.read_csv_rows(lines)
        assert len(rows) == 2
        assert rows[0] == ["A", "B", "C"]
        assert rows[1] == ["1", "2", "3"]

    def test_comma_delimiter(self):
        dm = DataManager()
        lines = ["A,B,C", "1,2,3"]
        rows = dm.read_csv_rows(lines)
        assert len(rows) == 2
        assert rows[0] == ["A", "B", "C"]

    def test_empty_lines(self):
        dm = DataManager()
        rows = dm.read_csv_rows([])
        assert rows == []


class TestParseMain:
    """Test per il parsing dell'orario classi."""

    def test_basic_parsing(self):
        dm = DataManager()
        csv_content = """LUNEDÌ;MARTEDÌ;MERCOLEDÌ;GIOVEDÌ;VENERDÌ;SABATO
CLASSE 1A
08:00;Rossi;Bianchi;Verdi;Neri;Gialli;Blu
09:00;Viola;Arancio;Rosa;Grigio;Marrone;Azzurro"""
        count = dm.parse_main(csv_content, is_path=False)
        assert count == 1
        assert "1A" in dm.db_main

    def test_empty_content(self):
        dm = DataManager()
        count = dm.parse_main("", is_path=False)
        assert count == 0
        assert dm.db_main == {}

    def test_multiple_classes(self):
        dm = DataManager()
        csv_content = """LUNEDÌ;MARTEDÌ;MERCOLEDÌ;GIOVEDÌ;VENERDÌ;SABATO
CLASSE 1A
08:00;Rossi;Bianchi;Verdi;Neri;Gialli;Blu
CLASSE 2B
08:00;Mario;Luigi;Peach;Toad;Yoshi;Bowser"""
        count = dm.parse_main(csv_content, is_path=False)
        assert count == 2
        assert "1A" in dm.db_main
        assert "2B" in dm.db_main

    def test_split_biennio_vendita_classes(self):
        dm = DataManager()
        csv_content = """Ora;LUNEDÌ;MARTEDÌ;MERCOLEDÌ;GIOVEDÌ;VENERDÌ;SABATO
CLASSE 1AB (PLESSO: CENTRALE)
08:00;Rossi;Bianchi;CANUPIETRO ORANI;Neri;Gialli;Blu
CLASSE 1AB (LAB SALA)
08:00;;;MURA;;;"""
        count = dm.parse_main(csv_content, is_path=False)
        assert count == 2
        assert "1ABIENNIO" in dm.db_main
        assert "1ABIENNIO VENDITA" in dm.db_main
        assert dm.db_main["1ABIENNIO"]["MERCOLEDÌ"][1]["text"] == "CANUPIETRO ORANI"
        assert dm.db_main["1ABIENNIO VENDITA"]["MERCOLEDÌ"][1]["text"] == "MURA"



class TestParseGeneric:
    """Test per il parsing del potenziamento e pagamento."""

    def test_basic_potenz(self):
        dm = DataManager()
        csv_content = "08:00;Rossi;09:00;Bianchi;10:00;Verdi;11:00;Neri;12:00;Gialli;13:00;Blu"
        loaded = dm.parse_generic(csv_content, key='pot', is_path=False)
        assert loaded > 0

    def test_empty_content(self):
        dm = DataManager()
        loaded = dm.parse_generic("", key='pot', is_path=False)
        assert loaded == 0


class TestGetContent:
    """Test per il metodo _get_content."""

    def test_string_input(self):
        dm = DataManager()
        lines, raw = dm._get_content("line1\nline2\nline3", is_path=False)
        assert len(lines) == 3
        assert lines[0] == "line1"

    def test_empty_string(self):
        dm = DataManager()
        lines, raw = dm._get_content("", is_path=False)
        assert lines == []  # Python's "".splitlines() returns []

    def test_nonexistent_file(self):
        dm = DataManager()
        lines, raw = dm._get_content("/nonexistent/path/file.csv", is_path=True)
        assert lines == []


class TestParseSos:
    """Test per il parsing dell'orario docenti di sostegno."""

    def test_parse_sos_from_workspace_file(self):
        dm = DataManager()
        file_path = "orario docenti di sostegno - 21sett.csv"
        if os.path.exists(file_path):
            loaded = dm.parse_sos(file_path, is_path=True)
            assert loaded == 234
            assert len(dm.db_sos) == 64
            # Verifica presenza docente Atzori Paola
            from gestione_sostituzioni.utils import to_slug
            slug = to_slug("ATZORI PAOLA")
            assert slug in dm.db_sos
            assert dm.db_sos[slug]['_name'] == "ATZORI PAOLA"
            assert "LUNEDÌ" in dm.db_sos[slug]
            assert 1 in dm.db_sos[slug]["LUNEDÌ"]
            assert dm.db_sos[slug]["LUNEDÌ"][1] == "1BABIENNIO"

    def test_parse_sos_string_content(self):
        from gestione_sostituzioni.utils import to_slug
        dm = DataManager()
        content = (
            "NOME DOCENTE,ORA,LUNEDI,MARTEDI,MERCOLEDI,GIOVEDI,VENERDI\n"
            "DOCENTE SOSTEGNO TEST,ORA,,,,,\n"
            ",1,SCIENZE 1B,,STORIA 5AS,,\n"
            ",2,TIC 1B,INGLESE 1B,ITALIANO 5AS,,\n"
        )
        loaded = dm.parse_sos(content, is_path=False)
        assert loaded == 5
        slug = to_slug("DOCENTE SOSTEGNO TEST")
        assert slug in dm.db_sos
        assert dm.db_sos[slug]["LUNEDÌ"][1] == "1BABIENNIO"
        assert dm.db_sos[slug]["MERCOLEDÌ"][1] == "5ASAL"


class TestExcelSupport:
    """Test per il caricamento di file Excel (.xlsx, .xls) da Excel 2000 in poi."""

    def test_read_xlsx_file(self, tmp_path):
        import openpyxl
        dm = DataManager()
        xlsx_file = str(tmp_path / "test_orario.xlsx")

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["NOME DOCENTE", "ORA", "LUNEDI", "MARTEDI", "MERCOLEDI", "GIOVEDI", "VENERDI"])
        ws.append(["DOCENTE EXCEL", "ORA", None, None, None, None, None])
        ws.append([None, 1, "SCIENZE 1B", None, "STORIA 5AS", None, None])
        wb.save(xlsx_file)

        loaded = dm.parse_sos(xlsx_file, is_path=True)
        assert loaded == 2
        from gestione_sostituzioni.utils import to_slug
        slug = to_slug("DOCENTE EXCEL")
        assert slug in dm.db_sos
        assert dm.db_sos[slug]["LUNEDÌ"][1] == "1BABIENNIO"
        assert dm.db_sos[slug]["MERCOLEDÌ"][1] == "5ASAL"

    def test_read_excel_function_direct(self, tmp_path):
        import openpyxl
        dm = DataManager()
        xlsx_file = str(tmp_path / "sample.xlsx")

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Col1", "Col2"])
        ws.append(["ValA", "ValB"])
        wb.save(xlsx_file)

        lines, raw = dm._read_excel(xlsx_file)
        assert len(lines) == 2
        assert "Col1;Col2" in lines[0]
        assert "ValA;ValB" in lines[1]


class TestParsePers:
    """Test per il parsing dell'orario personale docenti (Quadro 2)."""

    def test_parse_pers_horizontal_block(self):
        """Formato a blocchi orizzontali (es. Orario Professori Orizzontale)."""
        content = (
            "DOCENTE;Ora; Lunedì; Martedì; Mercoledì; Giovedì; Venerdì\n"
            "ALIAS;8:00; 4Beno c; 3Asal c; 2AB c;;\n"
            "ITALIANO;9:00; 4Bsal c; 3Asal c; 2AB c;;\n"
            ";10:00; R; 2AB c;; 4Bsal c; 4Beno c\n"
        )
        dm = DataManager()
        loaded = dm.parse_pers(content, is_path=False)
        assert loaded > 0
        assert "ALIAS" in dm.db_pers
        assert "LUNEDÌ" in dm.db_pers["ALIAS"]
        assert dm.db_pers["ALIAS"]["LUNEDÌ"][1] == "4BENO"
        assert dm.db_pers["ALIAS"]["LUNEDÌ"][2] == "4BSAL"
        # R non deve essere considerato come classe
        assert 3 not in dm.db_pers["ALIAS"]["LUNEDÌ"]
        # Martedì 1a ora
        assert dm.db_pers["ALIAS"]["MARTEDÌ"][1] == "3ASAL"

    def test_parse_pers_vertical_block(self):
        """Formato a blocchi verticali con prefisso DOCENTE [NOME]."""
        content = (
            "DOCENTE ALIAS;;;;;;\n"
            "Ora;Lunedì;Martedì;Mercoledì;Giovedì;Venerdì;Sabato\n"
            "08:00;2DB c;;3Beno c;;;\n"
            "09:00;2DB c;;1AB c;;;\n"
            "10:00;;;D;;;\n"
        )
        dm = DataManager()
        loaded = dm.parse_pers(content, is_path=False)
        assert loaded == 4
        assert "ALIAS" in dm.db_pers
        assert dm.db_pers["ALIAS"]["LUNEDÌ"][1] == "2DABIENNIO"
        assert dm.db_pers["ALIAS"]["MERCOLEDÌ"][1] == "3BENO"
        # 'D' (disposizione) non deve essere una classe
        assert "MERCOLEDÌ" not in dm.db_pers["ALIAS"] or 3 not in dm.db_pers["ALIAS"]["MERCOLEDÌ"]

    def test_parse_pers_legacy_single_row(self):
        """Formato legacy a riga singola (DOCENTE; 8:00; 1A; ...)."""
        content = (
            "DOCENTE;ORA_LUN;CLASSE_LUN;ORA_MAR;CLASSE_MAR\n"
            "ROSSI;8:00;1A;9:00;2B\n"
        )
        dm = DataManager()
        loaded = dm.parse_pers(content, is_path=False)
        assert loaded >= 1
        assert "ROSSI" in dm.db_pers

    def test_parse_pers_empty(self):
        dm = DataManager()
        loaded = dm.parse_pers("", is_path=False)
        assert loaded == 0
        assert dm.db_pers == {}



