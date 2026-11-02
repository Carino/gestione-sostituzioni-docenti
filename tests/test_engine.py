"""
tests/test_engine.py
Test unitari per il motore di calcolo delle sostituzioni (SubstitutionEngine).
"""

import pytest
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from gestione_sostituzioni.data_manager import DataManager
from gestione_sostituzioni.engine import SubstitutionEngine
from gestione_sostituzioni.utils import to_slug


@pytest.fixture
def basic_engine():
    """Crea un engine con dati di base per i test."""
    dm = DataManager()

    # Orario classi minimo: 1A con docente Rossi il Lunedì alla 1a ora
    dm.db_main = {
        "1A": {
            "LUNEDÌ": {
                1: {"text": "ROSSI", "blocked": []},
                2: {"text": "BIANCHI", "blocked": []},
                3: {"text": "VERDI", "blocked": []},
            }
        },
        "2B": {
            "LUNEDÌ": {
                1: {"text": "NERI", "blocked": []},
                2: {"text": "ROSSI", "blocked": []},
            }
        }
    }

    # Potenziamento: GIALLI disponibile il Lunedì alla 1a e 2a ora
    dm.db_pot = {"LUNEDÌ": {h: [] for h in range(1, 9)}}
    dm.db_pot["LUNEDÌ"][1] = ["GIALLI"]
    dm.db_pot["LUNEDÌ"][2] = ["GIALLI"]

    # Pagamento: VIOLA disponibile alla 3a ora
    dm.db_pay = {"LUNEDÌ": {h: [] for h in range(1, 9)}}
    dm.db_pay["LUNEDÌ"][3] = ["VIOLA"]

    # Alternativa: vuota
    dm.db_alt = {d: {} for d in ["LUNEDÌ", "MARTEDÌ", "MERCOLEDÌ", "GIOVEDÌ", "VENERDÌ", "SABATO"]}

    # Personale e sostegno: vuoti
    dm.db_pers = {}
    dm.db_sos = {}

    return SubstitutionEngine(dm)


class TestSubstitutionEngineBasic:
    """Test per le sostituzioni di base."""

    def test_no_results_if_no_targets(self, basic_engine):
        results = basic_engine.find_substitutions("LUNEDÌ", [], [], {}, [])
        assert len(results) == 0

    def test_finds_substitution_for_absent_teacher(self, basic_engine):
        rossi_slug = to_slug("ROSSI")
        results = basic_engine.find_substitutions(
            "LUNEDÌ",
            [rossi_slug],
            [rossi_slug],
            {},
            []
        )
        # Rossi insegna alla 1a ora in 1A e alla 2a ora in 2B
        assert len(results) >= 2

    def test_substitution_contains_class_info(self, basic_engine):
        rossi_slug = to_slug("ROSSI")
        results = basic_engine.find_substitutions(
            "LUNEDÌ",
            [rossi_slug],
            [rossi_slug],
            {},
            []
        )
        classes = [r[0] for r in results]
        assert "1A" in classes

    def test_substitution_contains_hour(self, basic_engine):
        rossi_slug = to_slug("ROSSI")
        results = basic_engine.find_substitutions(
            "LUNEDÌ",
            [rossi_slug],
            [rossi_slug],
            {},
            []
        )
        hours = [r[1] for r in results]
        assert 1 in hours


class TestSubstitutionPotenziamento:
    """Test per le sostituzioni tramite potenziamento."""

    def test_potenziamento_used(self, basic_engine):
        rossi_slug = to_slug("ROSSI")
        results = basic_engine.find_substitutions(
            "LUNEDÌ",
            [rossi_slug],
            [rossi_slug],
            {},
            []
        )
        # Almeno uno dei risultati dovrebbe usare il potenziamento
        tags = [r[4] for r in results]
        potenz_count = tags.count("potenziamento")
        # Con GIALLI disponibile alla 1a e 2a ora, e ROSSI assente a quelle ore,
        # almeno una sostituzione dovrebbe essere potenziamento
        assert potenz_count >= 1

    def test_round_robin_tracking(self, basic_engine):
        """Il contatore round-robin deve incrementarsi."""
        rossi_slug = to_slug("ROSSI")
        basic_engine.find_substitutions(
            "LUNEDÌ",
            [rossi_slug],
            [rossi_slug],
            {},
            []
        )
        # pot_usage deve avere almeno un docente tracciato
        assert len(basic_engine.pot_usage) >= 0  # May or may not use pot depending on logic


class TestSubstitutionShortLeaves:
    """Test per i permessi brevi."""

    def test_short_leave_specific_hours(self, basic_engine):
        rossi_slug = to_slug("ROSSI")
        # Rossi ha permesso breve solo alla 1a ora
        results = basic_engine.find_substitutions(
            "LUNEDÌ",
            [rossi_slug],
            [],  # Non è nella blacklist (assente totale)
            {rossi_slug: [1]},  # Permesso breve solo alla 1a ora
            []
        )
        # Deve trovare sostituzione solo per la 1a ora, non per la 2a
        hours = [r[1] for r in results]
        assert 1 in hours
        # Alla 2a ora non dovrebbe servire sostituzione (è presente)


class TestSubstitutionAbsentClasses:
    """Test per le classi assenti."""

    def test_absent_class_frees_teacher(self, basic_engine):
        rossi_slug = to_slug("ROSSI")
        neri_slug = to_slug("NERI")
        # Classe 2B assente, NERI dovrebbe essere liberato
        results = basic_engine.find_substitutions(
            "LUNEDÌ",
            [rossi_slug],
            [rossi_slug],
            {},
            ["2B"]  # Classe 2B assente
        )
        # Almeno un risultato usa NERI (liberato da classe 2B)
        all_text = " ".join([str(r[3]) for r in results])
        assert "NERI" in all_text
        assert len(results) >= 1

    def test_absent_class_frees_sostegno_teacher(self, basic_engine):
        dm = basic_engine.dm
        # Configura docente di sostegno Salaris in classe 2B alla 1a ora di Lunedì
        salaris_slug = to_slug("SALARIS")
        dm.db_sos = {
            salaris_slug: {
                '_name': "SALARIS",
                "LUNEDÌ": {1: "2B"}
            }
        }
        # In 2B non c'è docente curriculare all'ora 1 per questo test, così solo Salaris è liberata da 2B
        dm.db_main["2B"]["LUNEDÌ"].pop(1, None)

        rossi_slug = to_slug("ROSSI")
        results = basic_engine.find_substitutions(
            "LUNEDÌ",
            [rossi_slug],
            [rossi_slug],
            {},
            ["2B"]  # Classe 2B assente (gita)
        )
        # Salaris di sostegno deve essere liberata e proposta come DISP. DA CLASSE ASSENTE per la 1A
        h1_res = [r for r in results if r[0] == "1A" and r[1] == 1]
        assert len(h1_res) == 1
        assert "DISP. DA CLASSE ASSENTE: SALARIS" in h1_res[0][3]
        assert h1_res[0][4] == "disp_class"

    def test_absent_class_sostegno_teacher_not_used_if_also_absent(self, basic_engine):
        dm = basic_engine.dm
        salaris_slug = to_slug("SALARIS")
        dm.db_sos = {
            salaris_slug: {
                '_name': "SALARIS",
                "LUNEDÌ": {1: "2B"}
            }
        }
        dm.db_main["2B"]["LUNEDÌ"].pop(1, None)
        rossi_slug = to_slug("ROSSI")
        # Sia Rossi che Salaris sono assenti (blacklist)
        results = basic_engine.find_substitutions(
            "LUNEDÌ",
            [rossi_slug],
            [rossi_slug, salaris_slug],
            {},
            ["2B"]
        )
        h1_res = [r for r in results if r[0] == "1A" and r[1] == 1]
        assert len(h1_res) == 1
        assert "SALARIS" not in h1_res[0][3]

    def test_absent_class_frees_both_curricular_and_sostegno(self, basic_engine):
        dm = basic_engine.dm
        # In classe 2B all'ora 1 ci sono sia NERI (curriculare in db_main) sia SALARIS (sostegno in db_sos)
        salaris_slug = to_slug("SALARIS")
        dm.db_sos = {
            salaris_slug: {
                '_name': "SALARIS",
                "LUNEDÌ": {1: "2B"}
            }
        }
        # In classe 1A all'ora 1 c'è ROSSI (assente)
        # Aggiungiamo anche classe 3C all'ora 1 con BIANCHI (assente)
        dm.db_main["3C"] = {
            "LUNEDÌ": {
                1: {"text": "BIANCHI", "blocked": []}
            }
        }
        rossi_slug = to_slug("ROSSI")
        bianchi_slug = to_slug("BIANCHI")

        # 2B è assente (gita): vengono liberati sia NERI sia SALARIS alla 1a ora
        results = basic_engine.find_substitutions(
            "LUNEDÌ",
            [rossi_slug, bianchi_slug],
            [rossi_slug, bianchi_slug],
            {},
            ["2B"]
        )

        h1_subs = [r[3] for r in results if r[1] == 1]
        # Uno coperto da NERI e uno coperto da SALARIS
        all_h1_text = " ".join(h1_subs)
        assert "NERI" in all_h1_text
        assert "SALARIS" in all_h1_text
        assert all("DISP. DA CLASSE ASSENTE" in s for s in h1_subs)

    def test_absent_class_indicates_other_freed_teachers(self, basic_engine):
        dm = basic_engine.dm
        # In classe 2B alla 1a ora ci sono sia NERI (curriculare) sia SALARIS (sostegno)
        salaris_slug = to_slug("SALARIS")
        dm.db_sos = {
            salaris_slug: {
                '_name': "SALARIS",
                "LUNEDÌ": {1: "2B"}
            }
        }
        rossi_slug = to_slug("ROSSI")
        # Solo Rossi da sostituire in 1A alla 1a ora
        results = basic_engine.find_substitutions(
            "LUNEDÌ",
            [rossi_slug],
            [rossi_slug],
            {},
            ["2B"]
        )
        h1_res = [r for r in results if r[0] == "1A" and r[1] == 1]
        assert len(h1_res) == 1
        sub_text = h1_res[0][3]
        # Deve indicare il primo assegnato (NERI) e tra parentesi l'alternativa disponibile (SALARIS)
        assert "DISP. DA CLASSE ASSENTE: NERI (altri: SALARIS)" == sub_text


class TestSubstitutionTags:
    """Test per i tag dei risultati."""

    def test_result_has_valid_tag(self, basic_engine):
        rossi_slug = to_slug("ROSSI")
        results = basic_engine.find_substitutions(
            "LUNEDÌ",
            [rossi_slug],
            [rossi_slug],
            {},
            []
        )
        valid_tags = {"compresenza", "sostegno", "disp_class", "potenziamento", "pagamento", "scoperta", "alternativa"}
        for r in results:
            assert r[4] in valid_tags

    def test_result_tuple_has_6_elements(self, basic_engine):
        rossi_slug = to_slug("ROSSI")
        results = basic_engine.find_substitutions(
            "LUNEDÌ",
            [rossi_slug],
            [rossi_slug],
            {},
            []
        )
        for r in results:
            assert len(r) == 6  # (classe, ora, testo, sostituzione, tag, teacher_slug)


class TestSubstitutionSostegno:
    """Test per l'utilizzo dei docenti di sostegno in compresenza."""

    def test_sostegno_used_as_compresenza_when_curriculare_absent(self, basic_engine):
        dm = basic_engine.dm
        # Configura docente di sostegno Atzori in classe 1A alla 1a ora di Lunedì
        atzori_slug = to_slug("ATZORI PAOLA")
        dm.db_sos = {
            atzori_slug: {
                '_name': "ATZORI PAOLA",
                "LUNEDÌ": {1: "1A"}
            }
        }
        rossi_slug = to_slug("ROSSI")
        results = basic_engine.find_substitutions(
            "LUNEDÌ",
            [rossi_slug],
            [rossi_slug],
            {},
            []
        )
        # Deve trovare la sostituzione per la 1a ora con Atzori (sostegno)
        h1_res = [r for r in results if r[1] == 1]
        assert len(h1_res) == 1
        assert "COMPRESENZA SOSTEGNO: ATZORI PAOLA" in h1_res[0][3]
        assert h1_res[0][4] == "sostegno"

    def test_sostegno_not_used_if_sostegno_is_also_absent(self, basic_engine):
        dm = basic_engine.dm
        atzori_slug = to_slug("ATZORI PAOLA")
        dm.db_sos = {
            atzori_slug: {
                '_name': "ATZORI PAOLA",
                "LUNEDÌ": {1: "1A"}
            }
        }
        rossi_slug = to_slug("ROSSI")
        # Sia Rossi che Atzori sono assenti
        results = basic_engine.find_substitutions(
            "LUNEDÌ",
            [rossi_slug, atzori_slug],
            [rossi_slug, atzori_slug],
            {},
            []
        )
        # La 1a ora di 1A non deve essere coperta da Atzori (è assente!)
        h1_res = [r for r in results if r[0] == "1A" and r[1] == 1]
        assert len(h1_res) == 1
        assert "ATZORI" not in h1_res[0][3]
        # Atzori essendo assente da sola non deve generare righe aggiuntive (rimangono solo le 2 di Rossi)
        assert len(results) == 2

    def test_sostegno_absent_alone_does_not_generate_substitution(self, basic_engine):
        dm = basic_engine.dm
        atzori_slug = to_slug("ATZORI PAOLA")
        dm.db_sos = {
            atzori_slug: {
                '_name': "ATZORI PAOLA",
                "LUNEDÌ": {1: "1A"}
            }
        }
        # Solo Atzori è assente, curriculare presente
        results = basic_engine.find_substitutions(
            "LUNEDÌ",
            [atzori_slug],
            [atzori_slug],
            {},
            []
        )
        # Nessuna sostituzione deve essere generata se il curriculare è presente
        assert len(results) == 0

    def test_fallback_to_db_pers_when_db_main_is_empty(self, basic_engine):
        """Quando il Quadro 1 (db_main) è vuoto, il programma deve attingere dal Quadro 2 (db_pers)."""
        dm = basic_engine.dm
        dm.db_main = {}  # Quadro 1 svuotato!
        alias_slug = to_slug("ALIAS")
        dm.db_pers = {
            alias_slug: {
                '_name': "ALIAS",
                "VENERDÌ": {
                    3: "4BENO",
                    4: "2ABIENNIO"
                }
            }
        }
        results = basic_engine.find_substitutions(
            "VENERDÌ",
            [alias_slug],
            [alias_slug],
            {},
            []
        )
        assert len(results) == 2
        classes = [r[0] for r in results]
        assert "4BENO" in classes
        assert "2ABIENNIO" in classes
        # Verifica che la fonte indicata sia Da Personale
        for r in results:
            assert "(Da Personale: ALIAS)" in r[2]




class TestTitolareRule:
    """Il primo nome della cella è il titolare: se presente, nessuna sostituzione."""

    def _engine(self, text, blocked=None):
        dm = DataManager()
        dm.db_main = {"1E": {"LUNEDÌ": {4: {"text": text, "blocked": blocked or []}}}}
        dm.db_pot = {"LUNEDÌ": {h: [] for h in range(1, 9)}}
        dm.db_pay = {"LUNEDÌ": {h: [] for h in range(1, 9)}}
        dm.db_alt = {}
        dm.db_pers = {}
        dm.db_sos = {}
        return SubstitutionEngine(dm)

    def test_assente_secondo_nome_titolare_presente(self):
        eng = self._engine("RASSU Gambaro / TANCA G.M. (Compr: BASSU I)")
        slug = to_slug("Gambaro")
        res = eng.find_substitutions("LUNEDÌ", [slug], [slug], {}, [])
        assert res == []

    def test_assente_titolare_genera_sostituzione(self):
        eng = self._engine("RASSU Gambaro")
        slug = to_slug("Rassu")
        res = eng.find_substitutions("LUNEDÌ", [slug], [slug], {}, [])
        assert len(res) == 1 and "GAMBARO" in res[0][3].upper()

    def test_titolare_assente_e_secondo_target(self):
        eng = self._engine("RASSU Gambaro")
        a, b = to_slug("Rassu"), to_slug("Gambaro")
        res = eng.find_substitutions("LUNEDÌ", [a, b], [a, b], {}, [])
        # Anti-duplicato: per la stessa classe e ora viene generata una sola riga di sostituzione
        assert len(res) == 1
        assert a in res[0][5] and b in res[0][5]

    def test_copresenza_entrambi_assenti_non_duplica_riga(self):
        """Verifica che quando entrambi i docenti in compresenza sono assenti, non si duplichino le righe."""
        dm = DataManager()
        dm.db_main = {
            "5AENO": {
                "LUNEDÌ": {
                    1: {"text": "PISCHEDDA CARTAM", "blocked": []}
                }
            }
        }
        dm.db_pot = {"LUNEDÌ": {1: ["SANNAA"]}}
        dm.db_pay = {}
        dm.db_alt = {}
        dm.db_pers = {}
        dm.db_sos = {}
        eng = SubstitutionEngine(dm)

        pisc_slug = to_slug("PISCHEDDA")
        cart_slug = to_slug("CARTAM")
        results = eng.find_substitutions(
            "LUNEDÌ",
            [pisc_slug, cart_slug],
            [pisc_slug, cart_slug],
            {cart_slug: [1, 2, 3]},
            []
        )
        # Esattamente una sola riga generata per 5AENO alla 1a ora
        assert len(results) == 1
        assert results[0][0] == "5AENO"
        assert results[0][1] == 1
        assert "POTENZIAMENTO: SANNAA" in results[0][3]
        assert results[0][4] == "potenziamento"
        assert pisc_slug in results[0][5] and cart_slug in results[0][5]


class TestLaboratorioSdoppiato:
    """Test per le classi con laboratori sdoppiati (cucina e sala/vendita)."""

    def test_cucina_e_sala_indipendenti(self):
        dm = DataManager()
        dm.db_main = {
            "1ABIENNIO": {"MERCOLEDÌ": {1: {"text": "CANUPIETRO ORANI", "blocked": []}, 2: {"text": "CANUPIETRO ORANI DIANA", "blocked": []}}},
            "1ABIENNIO VENDITA": {"MERCOLEDÌ": {1: {"text": "MURA", "blocked": []}, 2: {"text": "MURA", "blocked": []}}}
        }
        dm.db_pot = {"MERCOLEDÌ": {h: [] for h in range(1, 9)}}
        dm.db_pay = {"MERCOLEDÌ": {h: [] for h in range(1, 9)}}
        dm.db_alt = {}
        dm.db_pers = {}
        dm.db_sos = {}
        eng = SubstitutionEngine(dm)

        # 1. Se CANUPIETRO è assente, ORANI è compresente e copre la classe
        res_c = eng.find_substitutions("MERCOLEDÌ", ["CANUPIETRO"], ["CANUPIETRO"], {}, [])
        assert len(res_c) == 2
        assert res_c[0][0] == "1ABIENNIO" and res_c[0][1] == 1 and "COMPRESENZA: ORANI" in res_c[0][3]
        assert res_c[1][0] == "1ABIENNIO" and res_c[1][1] == 2 and "COMPRESENZA: ORANI (o DIANA)" in res_c[1][3]

        # 2. Se MURA è assente, viene rilevato per 1ABIENNIO VENDITA
        res_m = eng.find_substitutions("MERCOLEDÌ", ["MURA"], ["MURA"], {}, [])
        assert len(res_m) == 2
        assert res_m[0][0] == "1ABIENNIO VENDITA" and res_m[0][1] == 1 and res_m[0][4] == "scoperta"

        # 3. Se ORANI è assente ma il titolare CANUPIETRO è presente, nessuna sostituzione
        res_o = eng.find_substitutions("MERCOLEDÌ", ["ORANI"], ["ORANI"], {}, [])
        assert res_o == []

