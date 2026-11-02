"""
gestione_sostituzioni.engine
Motore di calcolo delle sostituzioni.
"""

import re
import logging
from typing import Dict, List, Tuple, Any

from .data_manager import DataManager
from .utils import to_slug, is_subject, normalize_class

logger = logging.getLogger("GestioneSostituzioni.Engine")


class SubstitutionEngine:
    """Motore di calcolo per l'individuazione dei docenti sostituti."""

    def __init__(self, data_manager: DataManager):
        self.dm = data_manager
        self.pot_usage = {}  # Contatore d'uso per round-robin potenziamento
        self.last_freed_teachers = {h: [] for h in range(1, 9)}
        self.last_day = ""

    def find_substitutions(self, day: str, target_teachers_slugs: List[str],
                           blacklist: List[str], short_leaves: Dict[str, List[int]],
                           absent_classes_slugs: List[str]) -> List[Tuple[Any, ...]]:
        """
        Calcola le sostituzioni per i docenti indicati nel giorno specificato.

        Args:
            day: Giorno della settimana (es. "LUNEDÌ")
            target_teachers_slugs: Lista degli slug dei docenti da sostituire
            blacklist: Lista degli slug dei docenti assenti (non disponibili come sostituti)
            short_leaves: Dict slug -> lista ore di permesso breve
            absent_classes_slugs: Lista delle classi assenti (gite, ecc.)

        Returns:
            Lista di tuple (classe, ora, situazione, sostituzione, tag, teacher_slug)
        """
        results = []
        freed_teachers = {h: [] for h in range(1, 9)}
        used_teachers = {}
        self.last_day = day

        def cls_matches_absent(c_name: str, absent_list: List[str]) -> bool:
            """Verifica se il nome della classe corrisponde a una classe assente."""
            if not c_name or not absent_list:
                return False
            c_norm = normalize_class(c_name)
            for a in absent_list:
                if not a:
                    continue
                a_norm = normalize_class(a)
                if c_name == a or c_norm == a_norm:
                    return True
                if len(c_name) >= 2 and len(a) >= 2 and (c_name in a or a in c_name):
                    return True
                if len(c_norm) >= 2 and len(a_norm) >= 2 and (c_norm in a_norm or a_norm in c_norm):
                    return True
            return False

        # Raccolta docenti liberati dalle classi assenti
        if absent_classes_slugs:
            # 1. Docenti curriculari liberati (db_main)
            for cls_key in self.dm.db_main:
                if cls_matches_absent(cls_key, absent_classes_slugs) and day in self.dm.db_main[cls_key]:
                    for h, cell in self.dm.db_main[cls_key][day].items():
                        txt = cell['text']
                        clean_txt = re.sub(r'[\-\/\.;&,]', ' ', txt.upper())
                        toks = [w for w in clean_txt.split() if len(w) > 1 and "COMPR" not in w and not is_subject(w)]
                        for t in toks:
                            t_slug = to_slug(t)
                            is_absent_now = (t_slug in blacklist) or (t_slug in short_leaves and h in short_leaves[t_slug])
                            if not is_absent_now and not any(to_slug(x) == t_slug for x in freed_teachers[h]):
                                freed_teachers[h].append(t)

            # 2. Docenti di sostegno liberati dalle classi assenti (db_sos)
            for s_slug, s_data in self.dm.db_sos.items():
                if isinstance(s_data, dict) and day in s_data:
                    for h_val, sos_cls in s_data[day].items():
                        if h_val == '_name':
                            continue
                        try:
                            h_int = int(h_val)
                        except (ValueError, TypeError):
                            continue
                        if sos_cls and cls_matches_absent(sos_cls, absent_classes_slugs):
                            is_absent_now = (s_slug in blacklist) or (s_slug in short_leaves and h_int in short_leaves[s_slug])
                            if not is_absent_now and not any(to_slug(x) == s_slug for x in freed_teachers[h_int]):
                                s_name = s_data.get('_name', s_slug)
                                freed_teachers[h_int].append(s_name)

        self.last_freed_teachers = {h: list(freed_teachers[h]) for h in freed_teachers}

        def process_entry(cls, h, cell_data, source, teacher_slug):
            """Processa un singolo slot orario e trova il sostituto appropriato."""
            if teacher_slug in short_leaves:
                if h not in short_leaves[teacher_slug]:
                    return

            txt = cell_data['text']
            clean_txt = re.sub(r'[\-\/\.;&,]', ' ', txt.upper())
            tokens = [w for w in clean_txt.split() if len(w) > 1 and "COMPR" not in w and not is_subject(w)]

            if source == 'main':
                if not tokens:
                    return
                t_slug_tokens = [to_slug(x) for x in tokens]
                if teacher_slug not in t_slug_tokens:
                    return

                # Regola del titolare: il primo nome (prima di '(Compr:') è il docente
                # titolare dell'ora. Se è presente e l'assente è un altro nome della cella,
                # la classe è coperta: nessuna sostituzione. Non si applica alle celle di
                # classi sdoppiate (campo 'blocked'), dove ciascun docente segue il proprio gruppo.
                if not cell_data.get('blocked'):
                    main_part = re.split(r'\(\s*COMPR', clean_txt, maxsplit=1)[0]
                    main_tokens = [w for w in main_part.split() if len(w) > 1 and "COMPR" not in w and not is_subject(w)]
                    if main_tokens:
                        holder = to_slug(main_tokens[0])
                        if holder != teacher_slug:
                            holder_absent = (holder in blacklist) or (holder in short_leaves and h in short_leaves[holder])
                            if not holder_absent:
                                logger.debug(f"{cls} ora {h}: titolare {holder} presente, nessuna sostituzione per {teacher_slug}")
                                return

            res_txt = "CLASSE SCOPERTA"
            tag = "scoperta"
            subs = [t for t in tokens if to_slug(t) != teacher_slug] if source == 'main' else []

            valid_int = []
            for s in subs:
                s_slug = to_slug(s)
                blocked = any(to_slug(b) in s_slug for b in cell_data.get('blocked', []))
                is_sub_absent = (s_slug in blacklist) or (s_slug in short_leaves and h in short_leaves[s_slug])
                if not blocked and not is_sub_absent:
                    valid_int.append(s)

            if valid_int:
                res_txt = f"COMPRESENZA: {valid_int[0]}"
                if len(valid_int) > 1:
                    res_txt += f" (o {', '.join(valid_int[1:])})"
                tag = "compresenza"
            else:
                # Priorità 1.5: COMPRESENZA SOSTEGNO
                # Cerca se in questa classe, giorno e ora è presente un docente di sostegno
                valid_sos = []
                for s_slug, s_data in self.dm.db_sos.items():
                    if isinstance(s_data, dict) and day in s_data and h in s_data[day]:
                        sos_cls = s_data[day][h]
                        if sos_cls == cls or (len(sos_cls) >= 3 and (sos_cls in cls or cls in sos_cls)):
                            is_sos_absent = (s_slug in blacklist) or (s_slug in short_leaves and h in short_leaves[s_slug])
                            if not is_sos_absent:
                                sos_name = s_data.get('_name', s_slug)
                                valid_sos.append(sos_name)

                if valid_sos:
                    res_txt = f"COMPRESENZA SOSTEGNO: {valid_sos[0]}"
                    tag = "sostegno"
                else:
                    av_freed = [t for t in freed_teachers.get(h, []) if f"{h}-{to_slug(t)}" not in used_teachers and f"{h}-{t}" not in used_teachers]
                    if av_freed:
                        sel = av_freed[0]
                        sel_slug = to_slug(sel)
                        used_teachers[f"{h}-{sel_slug}"] = True
                        used_teachers[f"{h}-{sel}"] = True
                        res_txt = f"DISP. DA CLASSE ASSENTE: {sel}"
                        if len(av_freed) > 1:
                            res_txt += f" (altri: {', '.join(av_freed[1:])})"
                        tag = "disp_class"
                    else:
                        pots = self.dm.db_pot.get(day, {}).get(h, [])
                        valid_pots = []
                        for p in pots:
                            p_slug = to_slug(p)
                            is_p_absent = (p_slug in blacklist) or (p_slug in short_leaves and h in short_leaves[p_slug])
                            if not is_p_absent and f"{h}-{p_slug}" not in used_teachers:
                                valid_pots.append(p)

                        if valid_pots:
                            # Round-robin: scegli il docente con minor utilizzo nella sessione
                            sel = min(valid_pots, key=lambda p: self.pot_usage.get(to_slug(p), 0))
                            sel_slug = to_slug(sel)
                            self.pot_usage[sel_slug] = self.pot_usage.get(sel_slug, 0) + 1
                            used_teachers[f"{h}-{sel_slug}"] = True
                            res_txt = f"POTENZIAMENTO: {sel}"
                            tag = "potenziamento"
                            logger.debug(f"Potenziamento assegnato: {sel} (uso: {self.pot_usage[sel_slug]})")
                        else:
                            pays = self.dm.db_pay.get(day, {}).get(h, [])
                            valid_pays = []
                            for p in pays:
                                p_slug = to_slug(p)
                                is_p_absent = (p_slug in blacklist) or (p_slug in short_leaves and h in short_leaves[p_slug])
                                if not is_p_absent and f"{h}-{p_slug}" not in used_teachers:
                                    valid_pays.append(p)
                            if valid_pays:
                                sel_pay = valid_pays[0]
                                used_teachers[f"{h}-{to_slug(sel_pay)}"] = True
                                res_txt = f"PAGAMENTO: {sel_pay}"
                                if len(valid_pays) > 1:
                                    res_txt += f" (altri: {', '.join(valid_pays[1:])})"
                                tag = "pagamento"

            if source != 'alt':
                alt = self.dm.db_alt.get(day, {}).get(h, {}).get(cls)
                if alt:
                    a_slug = to_slug(alt)
                    is_a_absent = (a_slug in blacklist) or (a_slug in short_leaves and h in short_leaves[a_slug])
                    if not is_a_absent:
                        alt_note = f"ALTERNATIVA: {alt} (Non avvalentesi)"
                        if tag == "scoperta":
                            res_txt = f"CLASSE SCOPERTA\n{alt_note}"
                            tag = "alternativa"
                        else:
                            res_txt += f"\n{alt_note}"
            else:
                if tag == "scoperta":
                    res_txt = "CLASSE SCOPERTA (Non avvalentesi)"
                else:
                    res_txt += " (Non avvalentesi)"

            results.append((cls, h, txt, res_txt, tag, teacher_slug))

        # Fase 1: orario principale, personale, sostegno
        for teacher_slug in target_teachers_slugs:
            covered_hours = set()
            for cls in self.dm.db_main:
                if cls_matches_absent(cls, absent_classes_slugs):
                    continue
                if day in self.dm.db_main[cls]:
                    for h in range(1, 9):
                        if h in self.dm.db_main[cls][day]:
                            cell = self.dm.db_main[cls][day][h]
                            clean_txt = re.sub(r'[\-\/\.;&,]', ' ', cell['text'].upper())
                            tokens = [to_slug(w) for w in clean_txt.split() if len(w) > 1 and "COMPR" not in w and not is_subject(w)]
                            if teacher_slug in tokens:
                                covered_hours.add(h)
                                # Anti-duplicato per compresenza: se per questa classe e ora (nello stesso gruppo)
                                # è già stata generata una sostituzione da un altro compresente, unisci i docenti e salta
                                existing_idx = next((i for i, r in enumerate(results) if r[0] == cls and r[1] == h), None)
                                if existing_idx is not None and not cell.get('blocked'):
                                    prev_r = results[existing_idx]
                                    prev_slugs = str(prev_r[5]).split(',')
                                    if teacher_slug not in prev_slugs:
                                        combined_slugs = f"{prev_r[5]},{teacher_slug}"
                                        results[existing_idx] = (prev_r[0], prev_r[1], prev_r[2], prev_r[3], prev_r[4], combined_slugs)
                                    continue
                                process_entry(cls, h, cell, 'main', teacher_slug)

            if teacher_slug in self.dm.db_pers and day in self.dm.db_pers[teacher_slug]:
                for h_key, cls in self.dm.db_pers[teacher_slug][day].items():
                    if h_key == '_name':
                        continue
                    try:
                        h_int = int(h_key)
                    except (ValueError, TypeError):
                        continue
                    if h_int in covered_hours or cls_matches_absent(cls, absent_classes_slugs):
                        continue
                    covered_hours.add(h_int)
                    existing_idx = next((i for i, r in enumerate(results) if r[0] == cls and r[1] == h_int), None)
                    if existing_idx is not None:
                        prev_r = results[existing_idx]
                        prev_slugs = str(prev_r[5]).split(',')
                        if teacher_slug not in prev_slugs:
                            combined_slugs = f"{prev_r[5]},{teacher_slug}"
                            results[existing_idx] = (prev_r[0], prev_r[1], prev_r[2], prev_r[3], prev_r[4], combined_slugs)
                        continue
                    process_entry(cls, h_int, {'text': f'(Da Personale: {teacher_slug})', 'blocked': []}, 'pers', teacher_slug)

            # Nota: I docenti di sostegno assenti non generano sostituzione se il curricolare è presente.
            # Vengono utilizzati unicamente come compresenza a copertura del curricolare assente.

        # Fase 2: orario alternativa
        if day in self.dm.db_alt:
            for teacher_slug in target_teachers_slugs:
                for h in range(1, 9):
                    if h in self.dm.db_alt[day]:
                        for cls, alt_teacher in self.dm.db_alt[day][h].items():
                            if to_slug(alt_teacher) == teacher_slug:
                                if cls_matches_absent(cls, absent_classes_slugs):
                                    continue
                                if teacher_slug in short_leaves and h not in short_leaves[teacher_slug]:
                                    continue
                                # Se è già stata individuata una sostituzione per questa classe+ora, salta
                                if any(r[0] == cls and r[1] == h for r in results):
                                    continue
                                rel_note = " (Religione presente)" if (
                                    cls in self.dm.db_main and day in self.dm.db_main[cls] and h in self.dm.db_main[cls][day]
                                ) else ""
                                cell_info = {'text': f'MAT. ALTERNATIVA: {alt_teacher}{rel_note}', 'blocked': []}
                                process_entry(cls, h, cell_info, 'alt', teacher_slug)

        logger.info(f"Calcolo sostituzioni completato: {len(results)} risultati per {day}")
        return results
