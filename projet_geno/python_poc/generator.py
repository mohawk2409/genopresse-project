"""
Ce script a pour but de convertir un fichier JSON (contenant des données généalogiques
au format JSON-LD, probablement issues d'un fichier GEDCOM) en un rapport structuré
au format Microsoft Word (.docx).

Le processus se déroule en deux étapes principales :
1.  **Extraction et Transformation :** Le script lit le JSON d'entrée, le parcourt
    pour en extraire les informations pertinentes sur les individus, et transforme
    ces données brutes en un format plus lisible et structuré. Ce résultat
    intermédiaire est sauvegardé dans un fichier JSON de "rapport" (par exemple,
    extracted_report.json).
2.  **Génération du Document :** Le script lit ensuite ce fichier de rapport
    intermédiaire pour générer un document .docx bien formaté, avec des en-têtes,
    des pieds de page, et une présentation claire des informations pour chaque personne.
"""
import json
import re
import os
from pathlib import Path
import unicodedata
import sys
import string
sys.stdout.reconfigure(encoding='utf-8')
from typing import Any, Dict, List, Optional, Tuple
from datetime import date
import copy 
from docx import Document
from docx.text.paragraph import Paragraph
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, Cm
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.shared import RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_BREAK
from docx.enum.section import WD_SECTION
# --- Constantes ---
# Ces constantes définissent les chemins par défaut pour les fichiers d'entrée,
# de rapport intermédiaire et de sortie.
PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_JSON = PROJECT_ROOT / 'python_poc' / 'test.json'
DEFAULT_REPORT = PROJECT_ROOT / 'python_poc' / 'extracted_report.json'
DEFAULT_OUTPUT = PROJECT_ROOT / 'python_poc' / 'output.docx'
dictionnaire_jour = {0:"lundi",1:"mardi",2:"mercredi",3:"jeudi",4:"vendredi",5:"samedi",6:"dimanche"}
# Expression régulière pour identifier les chaînes de caractères qui sont des
# identifiants de référence dans le JSON (ex: _:A123).
ID_PATTERN = re.compile(r'^_:[A-Za-z0-9]+$')
SKIP_KEYS = {'MAP', 'CHAN', '_CREA', '_QUAL', 'QUAY', '_FNA'}
IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.tif'}
_FILE_SENTINEL = '__FILE__'
date_value = ""
people_by_id: Dict[str, "people"] = {}
# mentions_by_person[X] = ensemble des id de personnes dont la fiche mentionne X
# (comme conjoint, comme frère/sœur, comme parent) — sert à lister, dans l'index,
# toutes les pages où X apparaît, pas seulement sa propre fiche.
mentions_by_person: Dict[str, set] = {}
class people:

    def __init__(self):
        self.id = ""
        self.nom = []
        self.prenom = ""
        self.profession = ""
        self.lieu: Optional[List[tuple]] = None
        self.date = []
        self.paragraphe: Optional[Paragraph] = None
        self.generation: Optional[int]

    def setNom(self,nom):
        if nom and nom not in self.nom:
            self.nom.append(nom)

    def setPrenom(self,prenom):
        self.prenom = prenom

    def setProfession(self,profession):
        self.profession = profession

    def setLieu(self,lieu, raison, year):
        entry = (lieu, raison, year)
        if self.lieu == None:
            self.lieu = [entry]
        elif entry not in self.lieu:
            self.lieu.append(entry)

    def setDate(self,date):
        if date not in self.date:
            self.date.append(date)

    def setId(self,id):
        self.id = id

    def setParagraphe(self,paragraphe):
        self.paragraphe = paragraphe

    def getNom(self):
        return self.nom
    
    def getParagraphe(self):
        return self.paragraphe
    
    def __str__(self):
        return f"Nom: {self.nom}, Prenom: {self.prenom}, Lieu: {self.lieu}"


def _get_or_create_person(node: Dict[str, Any]) -> "people":
    """
    Point d'entrée unique pour créer/retrouver l'objet `people` d'un nœud JSON
    et y enregistrer son identité (id, prénom, nom, profession). Toujours le
    même objet pour un même '@id' (people_by_id.setdefault) : remplace les deux
    copies de cette logique qui existaient dans _display_name et
    _resolve_reference.
    """
    node_id = node.get('@id')
    p = people_by_id.setdefault(node_id, people())
    p.setId(node_id)
    name_value = node.get('foaf:name')
    if isinstance(name_value, dict):
        p.setPrenom(str(name_value.get('GIVN') or '').strip())
        p.setNom(str(name_value.get('SURN') or '').strip())
        p.setProfession(node.get('OCCU'))
    return p


def _record_lieu(person_id: Optional[str], event: Optional[Dict[str, Any]], raison: str, annee_suffixe: str = "") -> None:
    """
    Point d'entrée unique pour enregistrer le lieu d'un événement (Naissance,
    Mariage, Décès) sur l'objet `people` d'une personne. Calcule le lieu et
    l'année (repli sur "DNF" si la date est vide) de la même façon partout, au
    lieu de dupliquer ce calcul dans _resolve_reference, _person_biography et
    _build_person_summary.
    """
    if not person_id or not isinstance(event, dict):
        return
    place = _event_place(event)
    if not place:
        return
    date_str = str(event.get('dc:date', '')).strip()
    year = date_str.split()[-1] if date_str else "DNF"
    if annee_suffixe:
        year = f"{annee_suffixe} {year}"
    people_by_id.setdefault(person_id, people()).setLieu(place, raison, year)

# class dictionnaire_alphabet:
#     def __init__(self):
#         self.paragraphe: Optional[Paragraph] = None
#         self.letter_paragraphe: Optional[Paragraph] = None
# Cette variable prend les valeurs suivantes :
# 1 : le paragraphe de la lettre
# 2 : le paragraphe sous la lettre
# 3 : le paragraphe associé à un nom

dictionnaire_alphabet = {"A":[None,[],None,[]],"B":[None,[],None,[]],"C":[None,[],None,[]],"D":[None,[],None,[]],"E":[None,[],None,[]],"F":[None,[],None,[]],"G":[None,[],None,[]],
                         "H":[None,[],None,[]],"I":[None,[],None,[]],"J":[None,[],None,[]],"K":[None,[],None,[]],"L":[None,[],None,[]],"M":[None,[],None,[]],"N":[None,[],None,[]],"O":[None,[],None,[]],
                         "P":[None,[],None,[]],"Q":[None,[],None,[]],"R":[None,[],None,[]],"S":[None,[],None,[]],"T":[None,[],None,[]],"U":[None,[],None,[]],"V":[None,[],None,[]],"W":[None,[],None,[]],
                         "X":[None,[],None,[]],"Y":[None,[],None,[]],"Z":[None,[],None,[]]}

# --- Fonctions Utilitaires ("Helpers") ---
# Ces fonctions sont des "aides" internes au script, préfixées par `_`.

def _add_page_number(paragraph):
    """Ajoute les numéros de page (Page X of Y) à un paragraphe de pied de page."""
    paragraph.add_run('Page ')
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), 'PAGE \\* MERGEFORMAT')
    paragraph._p.append(fld)
    paragraph.add_run(' of ')
    fld2 = OxmlElement('w:fldSimple')
    fld2.set(qn('w:instr'), 'NUMPAGES \\* MERGEFORMAT')
    paragraph._p.append(fld2)


_bookmark_id_counter = 0


def _bookmark_name(identifiant: str) -> str:
    """
    Convertit un id JSON (ex: '_:731') en nom de signet Word valide : Word exige
    que ça commence par une lettre et ne contienne que lettres/chiffres/underscore
    (40 caractères max).
    """
    safe = re.sub(r'[^A-Za-z0-9_]', '_', str(identifiant))
    if not safe or not safe[0].isalpha():
        safe = 'bm_' + safe
    return safe[:40]


def _add_bookmark(paragraph, name: str) -> None:
    """
    Pose un signet sur ce paragraphe (sans texte visible), pour pouvoir y
    renvoyer ailleurs dans le document avec _add_page_ref.
    """
    global _bookmark_id_counter
    _bookmark_id_counter += 1
    start = OxmlElement('w:bookmarkStart')
    start.set(qn('w:id'), str(_bookmark_id_counter))
    start.set(qn('w:name'), name)
    end = OxmlElement('w:bookmarkEnd')
    end.set(qn('w:id'), str(_bookmark_id_counter))
    # w:pPr (propriétés du paragraphe), s'il existe, doit rester le tout premier
    # enfant de w:p : on insère le signet juste après, jamais avant.
    p = paragraph._p
    insert_at = 1 if len(p) and p[0].tag == qn('w:pPr') else 0
    p.insert(insert_at, start)
    p.append(end)


def _add_page_ref(paragraph, bookmark_name: str) -> None:
    """
    Insère un champ Word affichant le numéro de page du signet `bookmark_name`.
    Utilise la forme "complexe" (fldChar begin/separate/end) que Word génère
    lui-même pour ce type de champ — plus fiable que w:fldSimple pour PAGEREF.
    Un "?" de repli est affiché tant que le champ n'a pas été mis à jour par Word
    (à l'ouverture, ou via Ctrl+A puis F9).
    """
    run_begin = paragraph.add_run()
    fld_begin = OxmlElement('w:fldChar')
    fld_begin.set(qn('w:fldCharType'), 'begin')
    run_begin._r.append(fld_begin)

    run_instr = paragraph.add_run()
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = f' PAGEREF {bookmark_name} \\h '
    run_instr._r.append(instr)

    run_sep = paragraph.add_run()
    fld_sep = OxmlElement('w:fldChar')
    fld_sep.set(qn('w:fldCharType'), 'separate')
    run_sep._r.append(fld_sep)

    paragraph.add_run('?')

    run_end = paragraph.add_run()
    fld_end = OxmlElement('w:fldChar')
    fld_end.set(qn('w:fldCharType'), 'end')
    run_end._r.append(fld_end)


def _add_page_ref_if_different(paragraph, bookmark_name: str, previous_bookmark_name: str) -> None:
    """
    Insère ", <page de bookmark_name>" uniquement si cette page diffère de celle
    de `previous_bookmark_name`. Python ne connaît jamais le numéro de page réel
    (seul Word le calcule) : la comparaison se fait donc via un champ Word IF
    imbriquant deux PAGEREF, résolu par Word lui-même au moment du calcul des
    champs — pour éviter de répéter le même numéro deux fois de suite dans une
    liste de renvois (ex: "2, 2, 3" devient "2, 3").
    """
    def _instr(text):
        run = paragraph.add_run()
        node = OxmlElement('w:instrText')
        node.set(qn('xml:space'), 'preserve')
        node.text = text
        run._r.append(node)

    def _fld_char(char_type):
        run = paragraph.add_run()
        node = OxmlElement('w:fldChar')
        node.set(qn('w:fldCharType'), char_type)
        run._r.append(node)

    _fld_char('begin')  # champ IF englobant
    _instr(' IF ')
    _add_page_ref(paragraph, bookmark_name)
    _instr(' <> ')
    _add_page_ref(paragraph, previous_bookmark_name)
    _instr(' ", ')
    _add_page_ref(paragraph, bookmark_name)
    _instr('" "" ')
    _fld_char('separate')
    paragraph.add_run('')
    _fld_char('end')


def _add_page_refs(paragraph, owner_ids, order_by_id: Dict[str, int]) -> None:
    """
    Insère un renvoi de page pour chaque id de `owner_ids` : la propre fiche de
    la personne, plus chaque endroit où elle est mentionnée ailleurs (conjoint,
    fratrie, parent). Triés par ordre d'apparition dans le document
    (order_by_id). Deux pages consécutives identiques dans la liste ne sont
    affichées qu'une fois (voir _add_page_ref_if_different).
    """
    ids_sorted = sorted(owner_ids, key=lambda oid: order_by_id.get(oid, 0))
    if not ids_sorted:
        return
    _add_page_ref(paragraph, _bookmark_name(ids_sorted[0]))
    for previous_id, current_id in zip(ids_sorted, ids_sorted[1:]):
        _add_page_ref_if_different(paragraph, _bookmark_name(current_id), _bookmark_name(previous_id))


def _load_json(path: Path) -> Dict[str, Any]:
    """
    Charge un fichier JSON en essayant plusieurs encodages de caractères.
    Ceci est utile pour éviter les erreurs si le fichier n'est pas en UTF-8 standard.
    """
    for encoding in ['utf-8-sig', 'utf-16', 'utf-8', 'latin-1']:
        try:
            with open(path, 'r', encoding=encoding) as file_handle:
                return json.load(file_handle)
        except (UnicodeDecodeError, json.JSONDecodeError):
            continue
    raise ValueError(f"Could not decode {path} with any encoding")


def _save_json(path: Path, payload: Dict[str, Any]) -> None:
    """Sauvegarde un dictionnaire Python dans un fichier JSON avec l'encodage UTF-8."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, 'w', encoding='utf-8') as file_handle:
        json.dump(payload, file_handle, ensure_ascii=False, indent=2)


def _node_index(root: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    """
    Crée un "index" (annuaire) de tous les objets (nœuds) présents dans le JSON.
    Le dictionnaire résultant a pour clé l'ID de l'objet (`@id`) et pour valeur
    l'objet lui-même. Cela permet de retrouver très rapidement un objet par son ID.
    """
    index: Dict[str, Dict[str, Any]] = {}
    for node in root.get('@graph', []):
        if isinstance(node, dict) and '@id' in node:
            index[str(node['@id'])] = node
    return index


def _display_name(node: Dict[str, Any]) -> str:
    """
    Extrait et formate un nom complet et lisible pour une personne (nœud).
    Combine le prénom (GIVN) et le nom (SURN) ou utilise d'autres champs si indisponibles.
    """
    name_value = node.get('foaf:name')
    if isinstance(name_value, dict):
        given = str(name_value.get('GIVN') or '').strip()
        surname = str(name_value.get('SURN') or '').strip()
        if node.get('@id'):
            _get_or_create_person(node)
        gedvalue = str(name_value.get('GEDVALUE') or '').strip()
        if given and surname:
            return f"{given} {surname}"
        if gedvalue:
            return gedvalue.replace('/', ' ').replace('  ', ' ').strip()
    if isinstance(name_value, str) and name_value.strip():
        return name_value.strip().replace('/', ' ').replace('  ', ' ')
    return str(node.get('@id', 'Person'))

def _display_work(node: Dict[str, Any]) -> str:
    """
    Extrait et formate un nom complet et lisible pour une personne (nœud).
    Combine le prénom (GIVN) et le nom (SURN) ou utilise d'autres champs si indisponibles.
    """
    work_value = node.get('OCCU')
    if work_value:
        sentence = " " + _display_name(node) + " a été " + work_value + "."
        return sentence


def _gender_prefix(node: Dict[str, Any]) -> str:
    """Retourne un préfixe accordé en genre ("né", "née", "né(e)")."""
    gender = str(node.get('foaf:gender', '')).upper()
    if gender == 'F':
        return 'née'
    if gender == 'M':
        return 'né'
    return 'né(e)'


def _gender_role(node: Dict[str, Any]) -> str:
    """Retourne un rôle accordé en genre ("l'époux", "l'épouse")."""
    gender = str(node.get('foaf:gender', '')).upper()
    if gender == 'F':
        return 'l\'épouse'
    if gender == 'M':
        return 'l\'époux'
    return 'le conjoint'

def _gender_d(name: str) -> str:
    """Retourne un rôle accordé en genre ("l'époux", "l'épouse")."""
    if name[0] in ['A','E','I','O','U','Y']:
    # if name[0] == 'A' or name[0] == 'E' or name[0] == 'I' or name[0] == 'O' or name[0] == 'U' or name[0] == 'Y':
        return "d'"
    else:
        return 'de '


def _gender(node: Dict[str, Any]) -> str:
    """Retourne un rôle accordé en genre ("l'époux", "l'épouse")."""
    gender = str(node.get('foaf:gender', '')).upper()
    if gender == 'F':
        return 'fille'
    if gender == 'M':
        return 'fils'
    return 'enfant'

def _gender_marriage(node: Dict[str, Any]) -> str:
    """Retourne un rôle accordé en genre ("marié", "mariée", marié(e))."""
    gender = str(node.get('foaf:gender', '')).upper()
    if gender == 'F':
        return 'mariée'
    if gender == 'M':
        return 'marié'
    return 'marié(e)'

def _gender_death(node: Dict[str, Any]) -> str:
    """Retourne un rôle accordé en genre ("décédé", "décédée", décédé(e))."""
    gender = str(node.get('foaf:gender', '')).upper()
    if gender == 'F':
        return 'décédée'
    if gender == 'M':
        return 'décédé'
    return 'décédé(e)'

def _gender_pronoun(node: Dict[str, Any], majuscule=False) -> str:
    """Retourne un rôle accordé en genre ("l'époux", "l'épouse")."""
    gender = str(node.get('foaf:gender', '')).upper()
    if majuscule:
        if gender == 'F':
            return 'Elle'
        if gender == 'M':
            return 'Il'
    else:
        if gender == 'F':
            return 'elle'
        if gender == 'M':
            return 'il'
    return 'l\'enfant'


def indexationByName(tab):
    # indexation["nom"].sort()
    return tab

def _family_label(family_id: str, index: Dict[str, Dict[str, Any]], p) -> str:
    """
    Crée une étiquette lisible pour une famille, généralement en listant les conjoints.
    Par exemple : "Jean Dupont et Jeanne Durand".
    """
    family = index.get(str(family_id), {})
    if not family:
        return str(family_id)

    participants = family.get('bio:participant', [])
    if isinstance(participants, dict):
        participants = [participants]

    names: List[str] = []
    for participant in participants:
        if isinstance(participant, dict):
            participant_id = participant.get('@id')
            if participant_id and participant_id in index:
                display_name = _display_name(index[participant_id])
                names.append(display_name)

                parts = display_name.split()
                nom = parts[-1]
                prenom = " ".join(parts[:-1])

    if names:
        return ' et '.join(names)

    return str(family_id)


def _note_label(note: str, index: Dict[str, Dict[str, Any]]) -> str:
    """
    Ajoute les notes associées à la personne
    """
    note = index.get("NOTE")
    if not note:
        return "Aucune note associée."
    
    return note


def _reference_summary(reference_id: str, index: Dict[str, Dict[str, Any]],p) -> str:
    """
    Génère un résumé textuel pour un objet référencé par son ID.
    Par exemple, pour l'ID d'une personne, retourne son nom ; pour une famille, son label.
    """
    # target = index.get(str(reference_id), {})
    # if not target:
    #     return str(reference_id)

    # target_type = str(target.get('@type', ''))
    # if target_type == 'foaf:Person':
    #     final_texte = _display_name(target)
    #     marriage_phrase = _marriage_phrase(target, index)
    #     if marriage_phrase:
    #         final_texte += f", {marriage_phrase}"
    #     return final_texte
    # if target_type == 'bio:Relationship':
    #     return _family_label(str(reference_id), index,p)

    # title = target.get('TITL') or target.get('ABBR')
    # if title:
    #     return str(title)

    # name_value = target.get('foaf:name')
    # if isinstance(name_value, str):
    #     return name_value

    # if isinstance(name_value, dict):
    #     return _display_name(target)

    return str(reference_id)


def _person_marriages_in_order(person: Dict[str, Any], index: Dict[str, Dict[str, Any]],p) -> List[str]:
    """
    Retourne les ID des familles (bio:Relationship) où la personne est conjoint,
    triées par date de mariage croissante. Boucle sur 'bio:relationship' pour
    collecter les familles, les mariages sans date connue sont placés en dernier,
    dans leur ordre d'apparition.
    """
    relationships = person.get('bio:relationship', [])
    if isinstance(relationships, dict):
        relationships = [relationships]

    dated_families = []
    for order, rel in enumerate(relationships):
        if not (isinstance(rel, dict) and rel.get('@id') in index):
            continue
        fid = str(rel.get('@id'))
        family = index[fid]
        if family.get('@type') != 'bio:Relationship':
            continue

        # Les deux conjoints (participants) de cette famille : on ajoute leurs noms
        # de famille à l'index de la personne (utile pour retrouver, par ex., le nom
        # de jeune fille et le(s) nom(s) pris par mariage).
        participants = family.get('bio:participant', [])
        if isinstance(participants, dict):
            participants = [participants]
        for participant in participants[:2]:
            if not isinstance(participant, dict):
                continue
            participant_node = index.get(participant.get('@id'))
            if not participant_node:
                continue
            participant_name = participant_node.get('foaf:name')
            surname = str(participant_name.get('SURN') or '').strip() if isinstance(participant_name, dict) else ''
            gender = family.get("foaf:gender","")
            real_gender = participant_node.get('foaf:gender')
            print(real_gender)

            if surname and real_gender == "M":
                p.setNom(surname)

        marriage_event = family.get('bio:event')
        if isinstance(marriage_event, list):
            marriage_event = next(
                (e for e in marriage_event if isinstance(e, dict) and e.get('@type') == 'bio:Marriage'),
                None,
            )
        year = _event_year(marriage_event) if isinstance(marriage_event, dict) else None
        dated_families.append((fid, year, order))

    dated_families.sort(key=lambda item: (item[1] is None, item[1] if item[1] is not None else 0, item[2]))
    return [fid for fid, _, _ in dated_families]


def _marriage_phrase(person: Dict[str, Any], index: Dict[str, Dict[str, Any]]) -> str:
    """
    Construit une courte phrase indiquant que la personne s'est mariée, avec le nom
    du conjoint et, si connue, la date et le lieu du mariage (premier mariage connu).
    """
    p = people_by_id.setdefault(person.get('@id'), people())
    family_ids = _person_marriages_in_order(person, index, p)
    if not family_ids:
        return ''

    family = index.get(family_ids[0], {})
    participants = family.get('bio:participant', [])
    if isinstance(participants, dict):
        participants = [participants]
    person_id = person.get('@id')
    spouse_id = next(
        (pa.get('@id') for pa in participants
         if isinstance(pa, dict) and pa.get('@id') != person_id and pa.get('@id') in index),
        None,
    )
    if not spouse_id:
        return ''

    marriage_event = family.get('bio:event')
    if isinstance(marriage_event, list):
        marriage_event = next(
            (e for e in marriage_event if isinstance(e, dict) and e.get('@type') == 'bio:Marriage'),
            None,
        )

    marriage_phrase = ''
    if isinstance(marriage_event, dict):
        date_phrase = _event_date_phrase(str(marriage_event.get('dc:date', '')).strip())
        place_value = normaliser_adresse(_event_place(marriage_event))
        bits = [b for b in [date_phrase, f"à {place_value}" if place_value else ''] if b]
        marriage_phrase = ' '.join(bits)

    spouse_name = _display_name(index[spouse_id])
    verb_phrase = _gender_marriage(person)
    if marriage_phrase:
        return f"{verb_phrase} avec {spouse_name} {marriage_phrase}"
    return f"{verb_phrase} avec {spouse_name}"


def _resolve_reference(reference_id: str, index: Dict[str, Dict[str, Any]], current_node: Dict[str, Any],p) -> Dict[str, Any]:
    """
    Résout une référence ID en un objet structuré contenant des informations clés.
    Va chercher l'objet correspondant à l'ID dans l'index et retourne un dictionnaire
    avec son nom, son type (personne, famille), etc.
    """
    target = index.get(str(reference_id), {})
    resolved: Dict[str, Any] = {'@id': reference_id, 'display': _reference_summary(reference_id, index,p)}
    if isinstance(target.get('foaf:name'), dict):
        _get_or_create_person(target)

    if target.get('@type') == 'foaf:Person':
        resolved['type'] = 'person'
        resolved['name'] = _display_name(target)
        gender = target.get('foaf:gender')
        if gender:
            resolved['gender'] = gender
    elif target.get('@type') == 'bio:Relationship':
        resolved['type'] = 'family'
        resolved['name'] = _family_label(str(reference_id), index, p)
        participants = target.get('bio:participant', [])
        if isinstance(participants, dict):
            participants = [participants]
        participant_ids = [
            pa.get('@id') for pa in participants
            if isinstance(pa, dict) and pa.get('@id') in index
        ]
        resolved['participants'] = [_display_name(index[pid]) for pid in participant_ids]

        # Attributs 'date' et 'lieu' du mariage directement sur l'entrée bio:relationship,
        # pour un accès facile sans avoir à reparser le texte de 'display'.
        family_marriage_event = target.get('bio:event')
        if isinstance(family_marriage_event, list):
            family_marriage_event = next(
                (e for e in family_marriage_event if isinstance(e, dict) and e.get('@type') == 'bio:Marriage'),
                None,
            )
        if isinstance(family_marriage_event, dict):
            resolved['date'] = _pretty_date(str(family_marriage_event.get('dc:date', '')).strip())
            # p.setDate(resolved['date'])
            resolved['lieu'] = _event_place(family_marriage_event)
            # p.setLieu(resolved['lieu'])
        else:
            resolved['date'] = ''
            resolved['lieu'] = ''

        # Si la personne courante est l'un des conjoints de cette famille, on construit
        # une description de cette union précise : phrase de mariage (ou remariage selon
        # son rang chronologique parmi tous ses mariages), biographie du conjoint et
        # éventuels enfants de cette union. Une seule famille est décrite ici : celle
        # référencée par reference_id, pour éviter de dupliquer l'historique marital
        # complet dans chaque entrée de 'bio:relationship'.
        current_id = current_node.get('@id')
        if current_id in participant_ids:
            description_parts = []

            ordered_family_ids = _person_marriages_in_order(current_node, index,p)
            family_id = str(reference_id)
            rank = ordered_family_ids.index(family_id) if family_id in ordered_family_ids else 0

            spouse_id = next((pid for pid in participant_ids if pid != current_id), None)
            marriage_event = None
            if spouse_id:
                marriage_event = target.get('bio:event')
                if isinstance(marriage_event, list):
                    marriage_event = next(
                        (e for e in marriage_event if isinstance(e, dict) and e.get('@type') == 'bio:Marriage'),
                        None,
                    )
                marriage_phrase = ''
                age = None
                spouse_name = _display_name(index[spouse_id])
                if isinstance(marriage_event, dict):
                    date_phrase = _event_date_phrase(str(marriage_event.get('dc:date', '')).strip())
                    place_value = _event_place(marriage_event)
                    current_person = people_by_id.setdefault(current_id, people())
                    current_person.setId(current_id)
                    _record_lieu(current_id, marriage_event, "Mariage", spouse_name)
                    current_person.setDate(_pretty_date(str(marriage_event.get('dc:date', '')).strip()))
                    bits = [b for b in [date_phrase, f"à {place_value}" if place_value else ''] if b]
                    marriage_phrase = ' '.join(bits)
                    


                pronoun = _gender_pronoun(current_node, True)
                verb_phrase = _gender_marriage(current_node) if rank == 0 else f"re{_gender_marriage(current_node)}"
                age_phrase = f" à l'âge de {age} ans" if age is not None else ""
                if marriage_phrase:
                    description_parts.append(f"{pronoun} s'est {verb_phrase} avec {spouse_name} {marriage_phrase}{age_phrase}.")
                else:
                    description_parts.append(f"{pronoun} s'est {verb_phrase} avec {spouse_name}{age_phrase}.")
                description_parts.append(_person_biography(spouse_id, index))
                # if work_value and _display_work(current_node) not in description_parts:
                #     description_parts.append(work_value)

                children = target.get('CHIL')
                if children:
                    children_list = children if isinstance(children, list) else [children]
                    child_ids = [
                        c.get('@id') for c in children_list
                        if isinstance(c, dict) and c.get('@id') in index
                    ]
                    if child_ids:
                        intro = (
                            "De cette union est né un enfant :" if len(child_ids) == 1
                            else f"De cette union sont nés {len(child_ids)} enfants :"
                        )
                        description_parts.append(intro)
                        description_parts.extend(_person_biography(cid, index) for cid in child_ids)

            resolved['display'] = ' '.join(part for part in description_parts if part)
            birth_event = _find_event(current_node, 'bio:Birth')

            if isinstance(birth_event,dict):
                _record_lieu(current_id, birth_event, "Naissance")
                age = _compute_age(
                        str(birth_event.get('dc:date', '')).strip() if birth_event else '',
                        str(marriage_event.get('dc:date', '')).strip() if marriage_event else '',
                    )

            death_event = _find_event(current_node, 'bio:Death')
            _record_lieu(current_id, death_event, "Décès")

    elif target.get('@type') == '_:SOUR':
        # Gestion des sources (sources bibliographiques)
        resolved['type'] = 'source'
        resolved['type_source'] = target.get('TYPE', '')
        if target.get('WWW', '') != "":
            resolved['url'] = target.get('WWW', '')
        if target.get('DATA'):
            try:
                data_field = target.get('DATA')
                if isinstance(data_field, dict) and data_field.get('TEXT'):
                    resolved['text'] = str(data_field.get('TEXT', '')).strip()
            except Exception as source_error:
                print(f"WARNING: Error parsing source data - {source_error}")
        # Récupérer les pièces jointes (images, PDF, ...) attachées à la source
        source_media = target.get('OBJE')
        if source_media:
            resolved['OBJE'] = _pretty_nested_value('OBJE', source_media, index, current_node)
        resolved['display'] = resolved.get('display', f"Source: {resolved.get('type_source', 'inconnu')}")
    else:
        title = target.get('TITL')
        if title:
            resolved['title'] = title

    if current_node.get('@id') == reference_id:
        resolved['self'] = True

    return resolved


def _format_birth_event(event: Dict[str, Any], person_name: str, person_gender_prefix: str) -> str:
    """
    Formate une phrase descriptive pour un événement de naissance.

    Cette fonction prend un dictionnaire représentant un événement de naissance et
    construit une chaîne de caractères formatée qui inclut la date et le lieu
    de naissance.

    Exemple de sortie : "Jean Dupont est né le 01/01/1900 à Paris"

    Args:
        event: Le dictionnaire de l'événement, contenant potentiellement 'dc:date' et 'bio:place'.
        person_name: Le nom complet de la personne.
        person_gender_prefix: Le préfixe de genre ('né', 'née').

    Returns:
        Une chaîne de caractères formatée décrivant la naissance.
    """
    date_value = str(event.get('dc:date', '')).strip()
    weekday = _get_weekday(str(event.get('dc:date', '')).strip())
    
    place_info = event.get('bio:place')
    place_value = ''
    if isinstance(place_info, dict):
        # Si le lieu est un objet, on cherche la valeur dans 'GEDVALUE'
        place_value = _format_place(str(place_info.get('GEDVALUE', '')).strip())
    elif place_info:
        # Sinon, on utilise la valeur directement
        place_value = _format_place(str(place_info).strip())

    # Construire la phrase en plusieurs parties
    parts = [person_name, person_gender_prefix]
    if weekday and date_value:
        parts.append(f"le {weekday} {_pretty_date(date_value)}")
    elif date_value:
        parts.append(f"le {_pretty_date(date_value)}")
    if place_value:
        parts.append(f"à {place_value}")

    # Joindre les parties. 'est né' ou 'est née'
    if len(parts) > 2:
        return " ".join(parts)
    
    # S'il n'y a ni date ni lieu, on retourne une chaîne vide ou une info de base.
    return ""


def _collect_links(value: Any) -> List[str]:
    """Recursively collecte les URLs / chemins présents dans une valeur (event, dict, list, string)."""
    links: List[str] = []
    if isinstance(value, str):
        v = value.strip()
        if v.startswith('http') or v.startswith('/') or v.startswith('C') or re.search(r'\.(jpg|jpeg|png|gif|pdf|docx|txt|html|webp)$', v.lower()):
            links.append(v)
    elif isinstance(value, dict):
        for k, v in value.items():
            if k in {'WWW', 'URL', 'FILE', 'GEDVALUE'} and isinstance(v, (str, dict)):
                if isinstance(v, str):
                    links.extend(_collect_links(v))
                elif isinstance(v, dict) and 'GEDVALUE' in v and isinstance(v.get('GEDVALUE'), str):
                    links.extend(_collect_links(v.get('GEDVALUE')))
            else:
                links.extend(_collect_links(v))
    elif isinstance(value, list):
        for item in value:
            links.extend(_collect_links(item))

    # remove duplicates while preserving order
    seen = set()
    unique_links: List[str] = []
    for l in links:
        if l not in seen:
            seen.add(l)
            unique_links.append(l)
    return unique_links

def _pretty_date(date_string: str) -> str:
    """
    Formate une date au format GEDCOM (ex: "15 JAN 1950") 
    en format français lisible (ex: "15 janvier 1950").
    
    Args:
        date_string: Chaîne de date au format GEDCOM (JJ MMM AAAA)
        
    Returns:
        Date formatée en français ou la date originale si format invalide
    """
    if not date_string or not date_string.strip():
        return ''
    
    months = {
        "JAN": "janvier", "FEB": "février", "MAR": "mars", "APR": "avril", 
        "MAY": "mai", "JUN": "juin", "JUI": "juillet", "AUG": "août", 
        "SEP": "septembre", "OCT": "octobre", "NOV": "novembre", "DEC": "décembre"
    }
    
    try:
        parts = date_string.strip().split()
        if len(parts) < 2:
            return date_string  # Retourner la date telle quelle si format invalide
        
        day = parts[0]
        month_abbr = parts[1].upper()
        year = parts[2] if len(parts) > 2 else ''
        
        month_name = months.get(month_abbr, month_abbr)
        new_date = f"{day} {month_name} {year}".strip()
        return new_date
    except Exception as e:
        print(f"Warning: Could not parse date '{date_string}': {e}")
        return date_string


def _get_weekday(date_string: str) -> str:
    """
    Extrait le jour de la semaine a partir d'une date au format GEDCOM.
    
    Args:
        date_string: Chaine de date au format GEDCOM (JJ MMM AAAA, ex: "15 JAN 1950")
        
    Returns:
        Jour de la semaine en francais (ex: "lundi") ou chaine vide si parsing impossible
    """
    if not date_string or not date_string.strip():
        return ''
    
    # Dictionnaire mois GEDCOM -> numero
    months_dict = {
        "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
        "JUI": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12
    }
    
    try:
        parts = date_string.strip().split()
        if len(parts) < 3:
            return ''  # Format invalide
        
        day = int(parts[0])
        month_abbr = parts[1].upper()
        year = int(parts[2])
        
        month = months_dict.get(month_abbr)
        if month is None:
            return ''  # Mois invalide
        
        # Creer objet date Python
        date_obj = date(year, month, day)
        # Recuperer le jour de la semaine (0=lundi, 6=dimanche)
        weekday_num = date_obj.weekday()
        weekday_name = dictionnaire_jour.get(weekday_num, '')
        
        return weekday_name
    except Exception as e:
        print(f"Warning: Could not extract weekday from '{date_string}': {e}")
        return ''


def _find_event(node: Dict[str, Any], event_type: str) -> Optional[Dict[str, Any]]:
    """Retourne le premier événement d'un type donné (ex: 'bio:Birth') pour une personne."""
    events = node.get('bio:event', [])
    if not isinstance(events, list):
        events = [events]
    for event in events:
        if isinstance(event, dict) and event.get('@type') == event_type:
            return event
    return None


def _event_year(event: Optional[Dict[str, Any]]) -> Optional[int]:
    """Extrait l'année (4 chiffres) de la date d'un événement, si disponible."""
    if not event:
        return None
    match = re.search(r'(\d{4})', str(event.get('dc:date', '')))
    return int(match.group(1)) if match else None


def _date_parts(date_string: str) -> Optional[Tuple[int, Optional[int], Optional[int]]]:
    """
    Décompose une date GEDCOM en (année, mois, jour). Le mois et le jour valent
    None s'ils sont absents (ex: "1950" ou "JAN 1950"). Les préfixes GEDCOM
    d'imprécision (ABT, BEF, AFT, EST, CAL...) sont ignorés.

    Args:
        date_string: Chaîne de date au format GEDCOM (ex: "15 JAN 1950")

    Returns:
        Tuple (année, mois, jour) ou None si aucune année n'est identifiable.
    """
    if not date_string or not str(date_string).strip():
        return None

    months_dict = {
        "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
        "JUI": 7, "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12
    }

    year: Optional[int] = None
    month: Optional[int] = None
    day: Optional[int] = None

    for token in str(date_string).strip().split():
        token = token.strip('.,').upper()
        if token in months_dict:
            month = months_dict[token]
        elif token.isdigit():
            value = int(token)
            if len(token) == 4:
                year = value
            elif 1 <= value <= 31:
                day = value

    if year is None:
        return None
    if month is None:
        day = None
    return year, month, day


def _compute_age(start_date: str, end_date: str) -> Optional[int]:
    """
    Calcule l'âge (en années révolues) entre deux dates GEDCOM.

    L'anniversaire est pris en compte : si la date de fin précède l'anniversaire
    de l'année en cours, une année est retranchée. Quand le mois ou le jour
    manquent dans l'une des deux dates, le calcul se fait sur les seules années.

    Args:
        start_date: Date de départ au format GEDCOM (ex: naissance, "15 JAN 1950")
        end_date: Date de fin au format GEDCOM (ex: décès ou mariage)

    Returns:
        L'âge en années, ou None si les dates sont inexploitables ou incohérentes
        (âge négatif).
    """
    start = _date_parts(start_date)
    end = _date_parts(end_date)
    if not start or not end:
        return None

    start_year, start_month, start_day = start
    end_year, end_month, end_day = end

    age = end_year - start_year
    # Anniversaire pas encore passé à la date de fin : on retire une année.
    if None not in (start_month, end_month):
        if (end_month, end_day or 1) < (start_month, start_day or 1):
            age -= 1

    return age if age >= 0 else None


def _format_place(place_value: str) -> str:
    """Découpe le lieu par virgule, retire les segments vides (champs GEDCOM absents)
    et les espaces superflus, puis rejoint proprement avec ', '."""
    if not place_value:
        return ''
    parts = [part.strip() for part in place_value.split(',')]
    return ', '.join(part for part in parts if part)


def _event_place(event: Dict[str, Any]) -> str:
    """Extrait la valeur de lieu d'un événement (gère les formats objet et chaîne)."""
    place_info = event.get('bio:place')
    if isinstance(place_info, dict):
        return normaliser_adresse(_format_place(str(place_info.get('GEDVALUE', '')).strip()))
    if place_info:
        return normaliser_adresse(_format_place(str(place_info).strip()))
    return ''


def _event_date_phrase(date_value: str) -> str:
    """Formate une date GEDCOM en phrase ('le lundi 3 janvier 1938')."""
    if not date_value:
        return ''
    weekday = _get_weekday(date_value)
    pretty = _pretty_date(date_value).lower()
    return f"le {weekday} {pretty}" if weekday else (f"le {pretty}" if pretty else '')


def _person_biography(person_id: str, index: Dict[str, Dict[str, Any]]) -> str:
    """
    Construit une phrase biographique complète pour une personne : naissance (date, lieu)
    et, si connu, décès (date, lieu, âge atteint).
    """
    person = index.get(str(person_id), {})
    if not person:
        return ''

    name = _display_name(person)
    birth_event = _find_event(person, 'bio:Birth')
    death_event = _find_event(person, 'bio:Death')

    if birth_event:
        sentence = _format_birth_event(birth_event, name, _gender_prefix(person))
    else:
        sentence = ''
    if not sentence:
        sentence = f"{name} ({_gender_prefix(person)}, date de naissance inconnue)"
    sentence = sentence.strip().rstrip('.') + '.'

    if death_event:
        age = _compute_age(
            str(birth_event.get('dc:date', '')).strip() if birth_event else '',
            str(death_event.get('dc:date', '')).strip(),
        )
        death_verb = "est " + _gender_death(person)
        death_bits = [_gender_pronoun(person, True), death_verb]
        date_phrase = _event_date_phrase(str(death_event.get('dc:date', '')).strip())
        year = str(_event_year(death_event)) if str(_event_year(death_event)) != "" else "DNF"
        if date_phrase:
            death_bits.append(date_phrase)
        place_value = _event_place(death_event)
        if place_value:
            death_bits.append(f"à {place_value}")
        death_sentence = ' '.join(death_bits)
        if age is not None:
            death_sentence += f", à l'âge de {age} ans"
        sentence += f" {death_sentence}."

        _record_lieu(str(person_id),death_event,"Décès", "")

    return sentence


def _marriages_text(node: Dict[str, Any], index: Dict[str, Dict[str, Any]]) -> List[str]:
    """
    Pour chaque mariage de la personne (mariages et remariages, dans l'ordre
    chronologique) : la phrase de mariage ("s'est marié(e)"/"s'est remarié(e)" avec
    le/la conjoint(e), date, lieu), suivie de la biographie du/de la conjoint(e)
    (sa naissance puis, si connu, son décès). Ne renvoie que ce qui concerne le/la
    conjoint(e) — jamais les informations de la personne elle-même.
    """
    parts: List[str] = []
    person_id = node.get('@id')
    p = people_by_id.setdefault(person_id, people())
    family_ids = _person_marriages_in_order(node, index, p)

    for rank, fid in enumerate(family_ids):
        family = index.get(fid, {})
        participants = family.get('bio:participant', [])
        if isinstance(participants, dict):
            participants = [participants]
        spouse_id = next(
            (p.get('@id') for p in participants
             if isinstance(p, dict) and p.get('@id') != person_id and p.get('@id') in index),
            None,
        )
        if not spouse_id:
            continue
        mentions_by_person.setdefault(spouse_id, set()).add(person_id)

        marriage_event = family.get('bio:event')
        if isinstance(marriage_event, list):
            marriage_event = next(
                (e for e in marriage_event if isinstance(e, dict) and e.get('@type') == 'bio:Marriage'),
                None,
            )
        marriage_phrase = ''
        age = None
        if isinstance(marriage_event, dict):
            date_phrase = _event_date_phrase(str(marriage_event.get('dc:date', '')).strip())
            place_value = _event_place(marriage_event)
            bits = [b for b in [date_phrase, f"à {place_value}" if place_value else ''] if b]
            marriage_phrase = ' '.join(bits)
            birth_event = _find_event(node, 'bio:Birth')
            age = _compute_age(
                str(birth_event.get('dc:date', '')).strip() if birth_event else '',
                str(marriage_event.get('dc:date', '')).strip(),
            )

        pronoun = _gender_pronoun(node, True)
        verb_phrase = _gender_marriage(node) if rank == 0 else f"re{_gender_marriage(node)}"
        spouse_name = _display_name(index[spouse_id])
        sentence = f"{pronoun} s'est {verb_phrase} avec {spouse_name}"
        if marriage_phrase:
            sentence += f" {marriage_phrase}"
        if age is not None:
            sentence += f", à l'âge de {age} ans"
        parts.append(sentence + '.')

        spouse_bio = _person_biography(spouse_id, index)
        if spouse_bio:
            parts.append(spouse_bio)

    return parts


def _pretty_nested_value(key: str, value: Any, index: Dict[str, Dict[str, Any]], current_node: Dict[str, Any]) -> Any:
    """
    Fonction récursive au cœur de la transformation des données.
    Parcourt chaque champ d'un objet et le transforme en une valeur "jolie" et lisible.
    - Remplace les références ID par des résumés.
    - Formate les événements (naissance, mariage) en phrases descriptives.
    - Ajoute des labels 'display' pour un affichage direct.
    """
    p = people_by_id.setdefault(current_node.get('@id'), people())

    # Skip unwanted metadata keys globally
    if key in SKIP_KEYS:
        return None

    if key == 'foaf:name' and isinstance(value, dict):
        display_name = _display_name(current_node)
        target = current_node.get("GIVN")
        target = current_node.get("SURN")

        return {
            'display': display_name,
            **{sub_key: _pretty_nested_value(sub_key, sub_value, index, current_node) for sub_key, sub_value in value.items()},
        }

    if key == '_FIL' and isinstance(value, str):
        label = 'Enfant légitime' if value == 'LEGITIMATE_CHILD' else value
        return {'display': label, 'raw': value}

    if key in {'FAMC', 'FAMS'}:
        family_id = None
        if isinstance(value, dict):
            family_id = value.get('ofFamily')

        elif isinstance(value, str):
            family_id = value

        if isinstance(family_id, dict) and '@id' in family_id:
            family_id = family_id.get('@id')
        family_name = _family_label(str(family_id), index, p) if family_id else 'inconnue'
        role1 = _gender(current_node)
        pronoun = _gender_pronoun(current_node, True)
        prefix = _gender_prefix(current_node)
        
        # Extraire et formater la date depuis bio:event du noeud courant
        date_value = ''
        bio_events = current_node.get('bio:event', [])
        if not isinstance(bio_events, list):
            bio_events = [bio_events]
        
        # Chercher la première date disponible dans les événements
        for event in bio_events:
            if isinstance(event, dict) and 'dc:date' in event:
                date_value = str(event.get('dc:date', '')).strip()
                place_info = event.get('bio:place')
                # print("bio:place", place_info)
                place_value = ''
                if isinstance(place_info, dict):
                    # Si le lieu est un objet, on cherche la valeur dans 'GEDVALUE'
                    place_value = normaliser_adresse(_format_place(str(place_info.get('GEDVALUE', '')).strip()))
                elif place_info:
                    # Sinon, on utilise la valeur directement
                    place_value = normaliser_adresse(_format_place(str(place_info).strip()))
                break
        
        pretty_date = _pretty_date(date_value).lower() if date_value else ''
        weekday = _get_weekday(date_value) if date_value else ''
        
        # Construire la phrase avec jour de la semaine et date
        if weekday and pretty_date:
            date_info = f"{weekday} {pretty_date}"
        elif pretty_date:
            date_info = pretty_date
        else:
            date_info = ''
        
        display = f"{_display_name(current_node)}, {role1} {_gender_d(family_name)}{family_name}. {pronoun} est {prefix} le {date_info} à {place_value}."

        result: Dict[str, Any] = {'display': display}
        if family_id is not None:
            result['ofFamily'] = _resolve_reference(str(family_id), index, current_node,p)['display']
        if isinstance(value, dict):
            for sub_key, sub_value in value.items():
                if sub_key != 'ofFamily':
                    result[sub_key] = _pretty_nested_value(sub_key, sub_value, index, current_node)
        return result

    if key == 'bio:event':
        events = value if isinstance(value, list) else [value]
        pretty_events: List[Any] = []
        birth_date_value = ''
        for event in events:
            if not isinstance(event, dict):
                pretty_events.append(event)
                continue

            event_type = str(event.get('@type', ''))
            date_value = str(event.get('dc:date', '')).strip()
            place_info = event.get("bio:place")
            place_value = ''
            if isinstance(place_info, dict):
                # Si le lieu est un objet, on cherche la valeur dans 'GEDVALUE'
                place_value = normaliser_adresse(_format_place(str(place_info.get('GEDVALUE', '')).strip()))
            pretty_date_val = _pretty_date(date_value).lower() if date_value else ''
            weekday_val = _get_weekday(date_value) if date_value else ''
            
            # Construire l'info de date avec jour de la semaine
            if weekday_val and pretty_date_val:
                date_info = f"{weekday_val} {pretty_date_val}"
            elif pretty_date_val:
                date_info = pretty_date_val
            else:
                date_info = ''

            display = None
            if event_type == 'bio:Birth':
                display = _format_birth_event(event, _gender_pronoun(current_node, True), _gender_prefix(current_node))+"."
            elif event_type == 'bio:Marriage' and date_info:
                age = _compute_age(birth_date_value, date_value)
                age_text = f" à l'âge de {age} ans" if age is not None else ""
                display = f"{_display_name(current_node)} s'est {_gender_marriage(current_node)} le {date_info}{age_text}."
            elif event_type == 'bio:Death' and date_info:
                display = f"{_gender_pronoun(current_node,True)} est {_gender_death(current_node)} le {date_info} à {place_value}."
            elif date_info:
                display = f"Événement daté du {date_info}"

            note_value = str(event.get('NOTE', '')).strip()
            if display and note_value:
                display = f"{display} ({note_value})"

            pretty_event: Dict[str, Any] = {}
            # Ajout du type d'événement lisible
            if event_type:
                pretty_event['eventType'] = event_type

            # Collecte des liens / chemins présents dans l'événement
            links = _collect_links(event)
            if links:
                pretty_event['links'] = links

            if display:
                pretty_event['display'] = display
            for sub_key, sub_value in event.items():
                # On ignore les clés 'display' (que l'on vient de créer) et 'MAP'
                if sub_key not in {'display', 'MAP', '_QUAL', 'QUAY', '_FNA'}:
                    pretty_event[sub_key] = _pretty_nested_value(sub_key, sub_value, index, current_node)
            pretty_events.append(pretty_event)
            # print(pretty_event)

        return pretty_events if isinstance(value, list) else pretty_events[0]

    # if isinstance(value, str) and ID_PATTERN.match(value):
    #     return _resolve_reference(value, index, current_node)
    # if isinstance(value, dict) and '@id' in value and len(value) == 1:
    #     reference_id = str(value.get('@id'))
    #     return _resolve_reference(reference_id, index, current_node)

    if key == 'OBJE':
        entries = value if isinstance(value, list) else [value]
        pretty_media: List[Any] = []
        
        for idx, entry in enumerate(entries, 1):
            if not isinstance(entry, dict):
                pretty_media.append(entry)
                continue

            media_item: Dict[str, Any] = {}
            media_id = entry.get('@id')
            
            if media_id:
                media_item['@id'] = media_id
                target = index.get(str(media_id), {})
                
                # Extraire les informations de la pièce jointe
                title = target.get('TITL', '')
                file_value = ''
                file_info = target.get('FILE')
                
                if isinstance(file_info, dict):
                    file_value = str(file_info.get('GEDVALUE', '')).strip()
                elif isinstance(file_info, str):
                    file_value = file_info.strip()
                
                # Construire l'affichage
                display_bits = [f"Pièce jointe {idx}/{len(entries)}"]
                if title:
                    display_bits.append(f"Titre: {str(title).strip()}")
                if file_value:
                    display_bits.append(f"Fichier: {file_value}")
                
                # Ajouter le statut primaire si présent
                if entry.get('_PRIM') == 'YES':
                    display_bits.append("[PRIMAIRE]")
                
                media_item['display'] = ' | '.join(display_bits)
                media_item['title'] = title
                media_item['file'] = file_value
                
            if entry.get('_PRIM'):
                media_item['_PRIM'] = entry.get('_PRIM')
                
            pretty_media.append(media_item)
        
        return pretty_media if isinstance(value, list) else pretty_media[0]

    if isinstance(value, dict):
        ref_id = value.get('@id')
        if isinstance(ref_id, str) and ref_id in index:
            # Résout la référence (personne, famille ou source) en son résumé
            # structuré (nom, description, pièces jointes OBJE...).
            return _resolve_reference(ref_id, index, current_node,p)
        result: Dict[str, Any] = {}
        for sub_key, sub_value in value.items():
            if sub_key in SKIP_KEYS:
                continue
            sub_val = _pretty_nested_value(sub_key, sub_value, index, current_node)
            if sub_val is not None:
                result[sub_key] = sub_val
        return result

    if isinstance(value, list):
        if all(isinstance(item, (str, int, float, bool, type(None))) for item in value):
            return value if len(value) > 1 else (value[0] if value else '')
        processed: List[Any] = []
        for item in value:
            pv = _pretty_nested_value(key, item, index, current_node)
            if pv is None:
                continue
            processed.append(pv)
        # simplify single-item lists for readability
        if not processed:
            return []
        if len(processed) == 1:
            return processed[0]
        return processed

    return value

# TODO Faire les index :
# Récupérer chaque personne l'ajouter dans un tableau et rajouter à cette personne l'index associé 
# (si c'est la profession alors on ajoute la profession). 
# Pour afficher l'index il suffira de filtrer par soit la profession ou autre et de l'afficher avec le nom de la personne
# Ajouter les références

# Pour afficher les noms en gras, peut-être faire en sorte que c'est un autre "paragraphe" (avec le add_paragraphe) et tous les add_paragraphe des noms sont en gras. 


def _famc_parent_ids(node: Dict[str, Any], index: Dict[str, Dict[str, Any]]) -> List[str]:
    """
    Retourne les id des parents d'une personne, en remontant FAMC → famille →
    participants directement depuis les données brutes (indépendant du texte de
    filiation déjà généré).
    """
    person_id = node.get('@id')
    famc = node.get('FAMC')
    if famc is None:
        return []
    if not isinstance(famc, list):
        famc = [famc]

    parent_ids: List[str] = []
    for famc_entry in famc:
        family_id = None
        if isinstance(famc_entry, dict):
            of_family = famc_entry.get('ofFamily')
            if isinstance(of_family, dict):
                family_id = of_family.get('@id')
            elif isinstance(of_family, str):
                family_id = of_family
            elif isinstance(famc_entry.get('@id'), str):
                family_id = famc_entry.get('@id')
        elif isinstance(famc_entry, str):
            family_id = famc_entry

        if not family_id:
            continue

        family = index.get(str(family_id), {})
        participants = family.get('bio:participant', [])
        if isinstance(participants, dict):
            participants = [participants]

        for participant in participants:
            if not isinstance(participant, dict):
                continue
            parent_id = participant.get('@id')
            if not parent_id or parent_id == person_id or parent_id not in index:
                continue
            parent_ids.append(parent_id)

    return parent_ids


def _build_person_summary(node: Dict[str, Any], person_dict: Dict[str, Any], index: Dict[str, Dict[str, Any]]) -> str:
    """
    Rassemble en un seul texte continu les informations d'une personne qui, sinon,
    seraient dispersées entre plusieurs champs (FAMC, bio:event, OCCU, bio:relationship,
    rel:siblingOf), dans cet ordre précis :
    1. Filiation ("fils/fille de ...")
    2. Naissance de la personne
    3. Mariage(s)/remariage(s) : phrase de mariage puis naissance et décès du/de la
       conjoint(e) (jamais ceux de la personne elle-même)
    4. Autres événements de vie (résidence, etc.)
    5. Profession de la personne
    6. Décès de la personne (toujours en dernier)
    7. Fratrie
    """
    parts: List[str] = []
    node_id = node.get('@id')

    # 1. Filiation — inchangé
    famc = person_dict.get('FAMC')
    if isinstance(famc, dict) and famc.get('display'):
        filiation = str(famc['display']).split('.', 1)[0].strip()
        if filiation:
            parts.append(filiation + '.')
    if node_id:
        for parent_id in _famc_parent_ids(node, index):
            mentions_by_person.setdefault(parent_id, set()).add(node_id)

    events = person_dict.get('bio:event', [])
    if isinstance(events, dict):
        events = [events]
    events = [e for e in events if isinstance(e, dict) and e.get('display')]

    birth_display = next((e['display'] for e in events if e.get('eventType') == 'bio:Birth'), None)
    death_event = next((e for e in events if e.get('eventType') == 'bio:Death'), None)
    death_display = death_event.get('display') if death_event else None
    death_date = death_event.get('dc:date') if death_event else None
    death_place_info = death_event.get('bio:place') if death_event else None
    death_place = normaliser_adresse(death_place_info.get('GEDVALUE')) if isinstance(death_place_info, dict) else death_place_info
    _record_lieu(node.get('@id'), death_event, "Décès")
    other_displays = [e['display'] for e in events if e.get('eventType') not in {'bio:Birth', 'bio:Death','bio:relationship'}]

    # 2. Naissance de la personne — inchangé
    if birth_display:
        parts.append(str(birth_display).strip())

    # 3. Mariage(s)/remariage(s) : conjoint(e), jamais la personne elle-même
    parts.extend(_marriages_text(node, index))

    # 4. Autres événements de vie (résidence, etc.)
    parts.extend(str(d).strip() for d in other_displays)

    # 5. Profession de la personne
    work_value = _display_work(node)
    if work_value:
        parts.append(work_value.strip())

    # 6. Décès de la personne — toujours en dernier
    if death_display:
        # parts.append("[")
        parts.append(str(death_display).strip())
        # parts.append(str(death_place).strip())
        # parts.append(str(death_date).strip())
        # parts.append("]")

    # 7. Fratrie
    siblings = person_dict.get('rel:siblingOf')
    if siblings:
        siblings_list = siblings if isinstance(siblings, list) else [siblings]
        # 'name' d'abord : _reference_summary ne renvoyant que l'identifiant,
        # 'display' vaut "_:7311" alors que 'name' porte le vrai nom affichable.
        names = [
            s.get('name') or s.get('display') for s in siblings_list
            if isinstance(s, dict) and (s.get('name') or s.get('display'))
        ]
        if names:
            intro = "Frère ou sœur :" if len(names) == 1 else "Frères et sœurs :"
            parts.append(f"{intro} {', '.join(str(n) for n in names)}.")
        if node_id:
            for s in siblings_list:
                sibling_id = s.get('@id') if isinstance(s, dict) else None
                if sibling_id:
                    mentions_by_person.setdefault(sibling_id, set()).add(node_id)

    return ' '.join(pa for pa in parts if pa)


def _all_person_names(index: Dict[str, Dict[str, Any]]) -> List[str]:
    """
    Tous les noms affichables des personnes du fichier, du plus long au plus
    court. L'ordre importe pour le découpage des résumés : on veut reconnaître
    "Jeannine HRYSIO" en entier plutôt que le seul "HRYSIO" qu'il contient.
    """
    noms = set()
    for node in index.values():
        if not isinstance(node, dict) or node.get('@type') != 'foaf:Person':
            continue
        nom = _display_name(node)
        # _display_name retombe sur l'@id quand la personne n'a pas de nom :
        # ces identifiants n'ont rien à faire en gras dans le texte.
        if nom and not ID_PATTERN.match(nom):
            noms.add(nom)
    return sorted(noms, key=len, reverse=True)


def _split_summary_by_names(summary: str, noms: List[str]) -> Tuple[List[str], List[str]]:
    """
    Scinde le résumé à chaque nom de personne rencontré, pour pouvoir mettre les
    noms en gras à l'affichage (la mise en forme dans un .docx s'applique à un
    "run", donc un nom en gras doit être un morceau de texte séparé).

    Retourne (morceaux, noms_trouvés) tels que
    morceaux[0] + noms_trouvés[0] + morceaux[1] + noms_trouvés[1] + ...
    reconstitue exactement `summary`. Il y a toujours un morceau de plus que de
    noms (le dernier morceau peut être vide).
    """
    if not summary:
        return [''], []
    if not noms:
        return [summary], []

    motif = re.compile('|'.join(re.escape(nom) for nom in noms))
    morceaux: List[str] = []
    noms_trouves: List[str] = []
    position = 0
    for trouve in motif.finditer(summary):
        morceaux.append(summary[position:trouve.start()])
        noms_trouves.append(trouve.group())
        position = trouve.end()
    morceaux.append(summary[position:])
    return morceaux, noms_trouves


def extract_report(json_path: Optional[str] = None, report_path: Optional[str] = None) -> Dict[str, Any]:
    """
    Orchestre la première étape : l'extraction et la transformation des données.
    1. Charge le JSON d'entrée.
    2. Crée un index des objets pour des recherches rapides.
    3. Itère sur chaque personne, transforme ses données avec `_pretty_nested_value`.
    4. Sauvegarde le résultat dans un fichier de rapport intermédiaire (JSON).
    """
    json_file = Path(json_path) if json_path else DEFAULT_JSON
    report_file = Path(report_path) if report_path else DEFAULT_REPORT

    root = _load_json(json_file)
    index = _node_index(root)

    generation_map, bounds_by_gen = _compute_generations(root, index)

    report: Dict[str, Any] = {
        'sourceFile': str(json_file),
        'people': []
    }

    print("Début du traitement du fichier JSON...")
    graph = root.get('@graph', [])
    print(f"Nombre total d'éléments dans '@graph': {len(graph)}")

    # Calculés une seule fois : servent à scinder chaque résumé à chaque nom.
    noms_personnes = _all_person_names(index)

    for order, node in enumerate(graph, start=1):
        try:
            if not isinstance(node, dict) or node.get('@type') != 'foaf:Person':
                continue

            node_id = node.get('@id', 'ID inconnu')
            generation = generation_map.get(str(node_id))
            generation_number = generation[0] if isinstance(generation, tuple) else generation
            bound = bounds_by_gen.get(generation_number)

            print(f"Traitement de la personne {order}/{len(graph)} : {node_id} - {_display_name(node)}"
                  + (f" [Génération {generation}]" if generation else ""))

            heading = _display_name(node)

            person_dict: Dict[str, Any] = {}
            for key, value in node.items():
                # Ne pas inclure ces métadonnées de création/modification
                if key in {'@id', '@type', '_CREA', 'CHAN', 'MAP', '_QUAL', 'QUAY', '_FNA'}:
                    continue
                person_dict[key] = _pretty_nested_value(key, value, index, node)

            summary = _build_person_summary(node, person_dict, index)
            # 'summary' est découpé à chaque nom, et 'nom' contient ces noms dans
            # le même ordre : à l'affichage, on alterne un morceau puis un nom en
            # gras (voir generate_docx).
            morceaux_summary, noms_summary = _split_summary_by_names(summary, noms_personnes)

            report['people'].append(
                {
                    'order': order,
                    'id': node.get('@id'),
                    'type': node.get('@type'),
                    'title': str(heading),
                    'generation': generation,
                    'bound': bound,
                    'summary': morceaux_summary,
                    'nom': noms_summary,
                    'lines': person_dict,
                }
            )
        except Exception as e:
            print(f"--- ERREUR ---")
            print(f"Une erreur est survenue lors du traitement de l'élément {order} (ID: {node.get('@id', 'inconnu')})")
            print(f"Erreur : {e}")
            print(f"Données de l'élément problématique : {node}")
            print(f"----------------")
            # On continue avec l'élément suivant
            continue

    print(f"Fin du traitement. {len(report['people'])} personnes extraites.")
    _save_json(report_file, report)
    return report


def _set_header_footer(document: Document, title: str) -> None:
    """Configure l'en-tête et le pied de page du document Word."""
    section = document.sections[0]

    header = section.header
    header_paragraph = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    header_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    header_run = header_paragraph.add_run(title)
    header_run.bold = True
    header_run.font.size = Pt(10)

    footer = section.footer
    footer_paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    footer_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_page_number(footer_paragraph)

def _compute_generations(root: Dict[str, Any], index: Dict[str, Dict[str, Any]]) -> Tuple[Dict[str, Any], Dict[int, List[int]]]:
    """
    Calcule la génération de chaque personne à partir du premier foaf:Person du graph.
    La personne de référence est génération 1, ses parents génération 2, etc.
    Utilise un BFS (parcours en largeur) en remontant via FAMC → famille → parents (FAMS).
    """
    # Trouver la personne de référence (premier foaf:Person)
    ref_id = None
    for node in root.get('@graph', []):
        if isinstance(node, dict) and node.get('@type') == 'foaf:Person':
            ref_id = node.get('@id')
            break

    if not ref_id:
        return {}, {}

    generations: Dict[str, (int, int)] = {ref_id: 1}
    queue = [(ref_id, 1)]
    # Plage [année min, année max] des naissances/décès connus, par génération
    # (et non toutes générations confondues) : bounds_by_gen[3] couvre uniquement
    # les personnes de génération 3.
    bounds_by_gen: Dict[int, List[int]] = {}

    def _record_year(gen: int, year: Optional[int]) -> None:
        if year is None:
            return
        if gen not in bounds_by_gen:
            bounds_by_gen[gen] = [year, year]
        else:
            bounds_by_gen[gen][0] = min(bounds_by_gen[gen][0], year)
            bounds_by_gen[gen][1] = max(bounds_by_gen[gen][1], year)

    while queue:
        person_id, gen = queue.pop(0)
        person = index.get(person_id, {})

        birth_year = _event_year(_find_event(person, 'bio:Birth'))
        death_year = _event_year(_find_event(person, 'bio:Death'))
        _record_year(gen, birth_year)
        _record_year(gen, death_year)

        # Remonter vers les parents via FAMC (famille où la personne est enfant)
        famc = person.get('FAMC')
        if famc is None:
            continue
        if not isinstance(famc, list):
            famc = [famc]

        for famc_entry in famc:
            # Extraire l'ID de la famille
            family_id = None
            if isinstance(famc_entry, dict):
                of_family = famc_entry.get('ofFamily')
                if isinstance(of_family, dict):
                    family_id = of_family.get('@id')
                elif isinstance(of_family, str):
                    family_id = of_family
                elif isinstance(famc_entry.get('@id'), str):
                    family_id = famc_entry.get('@id')
            elif isinstance(famc_entry, str):
                family_id = famc_entry

            if not family_id:
                continue

            family = index.get(str(family_id), {})
            participants = family.get('bio:participant', [])
            if isinstance(participants, dict):
                participants = [participants]

            for participant in participants:
                if not isinstance(participant, dict):
                    continue
                parent_id = participant.get('@id')
                if not parent_id or parent_id == person_id:
                    continue
                parent = index.get(str(parent_id), {})
                if parent.get('@type') != 'foaf:Person':
                    continue
                if parent_id not in generations:
                    generations[parent_id] = (gen+1,len(participants))
                    # generations[parent_id] = gen + 1
                    queue.append((parent_id, gen + 1))

    return generations, bounds_by_gen


def _dict_to_text_lines(data: Any, indent: int = 0) -> List[str]:
    """Recursively convert a dict/value structure back to indented text lines for display."""
    prefix = '  ' * indent
    lines: List[str] = []
    
    if isinstance(data, dict):
        if 'display' in data:
            lines.append(f"{prefix}{data['display']}")
        for key, value in data.items():
            if key == 'display' or key in SKIP_KEYS:
                continue
            # Afficher explicitement les liens si présents
            if key == 'links':
                if isinstance(value, list):
                    for l in value:
                        lines.append(f"{prefix}  → Lien: {l}")
                else:
                    lines.append(f"{prefix}  → Lien: {value}")
                continue
            # Gérer les pièces jointes de manière spéciale
            if key == 'OBJE':
                lines.append(f"{prefix}📎 PIÈCES JOINTES:")
                entries = value if isinstance(value, list) else [value]
                for item in entries:
                    if isinstance(item, dict):
                        if 'display' in item:
                            lines.append(f"{prefix}  {item['display']}")
                        file_path = item.get('file', '')
                        if file_path:
                            lines.append(f"{_FILE_SENTINEL}{file_path}")
                    else:
                        lines.append(f"{prefix}  {item}")
                continue
            if isinstance(value, dict):
                lines.append(f"{prefix}{key}:")
                lines.extend(_dict_to_text_lines(value, indent + 1))
            elif isinstance(value, list):
                if not value:
                    lines.append(f"{prefix}{key}: []")
                elif all(isinstance(v, (str, int, float, bool, type(None))) for v in value):
                    # Simple list
                    lines.append(f"{prefix}{key}: {', '.join(str(v) for v in value)}")
                else:
                    # Complex list
                    lines.append(f"{prefix}{key}:")
                    for idx, item in enumerate(value, 1):
                        if isinstance(item, dict):
                            lines.append(f"{prefix}  [{idx}]")
                            lines.extend(_dict_to_text_lines(item, indent + 2))
                        else:
                            lines.append(f"{prefix}  [{idx}] {item}")
            else:
                lines.append(f"{prefix}{key}: {value}")
    elif isinstance(data, list):
        for idx, item in enumerate(data, 1):
            if isinstance(item, dict):
                lines.append(f"{prefix}[{idx}]")
                lines.extend(_dict_to_text_lines(item, indent + 1))
            else:
                lines.append(f"{prefix}[{idx}] {item}")
    else:
        lines.append(f"{prefix}{data}")
    
    return lines


def _create_cover_page(document: Document, first_person_name: str):
    """
    Crée une page de garde simple au début du document.
    """
    # Ajoute le nom, centré
    p_nom = document.add_paragraph()
    p_nom.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_nom = p_nom.add_run("Nom\n") # Remplacez "Nom" par le texte souhaité
    run_nom.font.size = Pt(24)
    run_nom.bold = True

    # Ajoute le titre, centré
    p_titre = document.add_paragraph()
    p_titre.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_titre = p_titre.add_run("TITRE\n") # Remplacez "TITRE" par le texte souhaité
    run_titre.font.size = Pt(48)
    run_titre.font.name = "Georgia"
    run_titre.italic = True

    # Ajoute le nom de la première personne
    font = document .styles['Normal'].font
    font.name = "Garamond"
    p_personne = document.add_paragraph()
    p_personne.alignment = WD_ALIGN_PARAGRAPH.CENTER
    # p_personne.font = "Garamond"
    run_personne = p_personne.add_run(f"de {first_person_name}\n\n")
    run_personne.font.size = Pt(18)
    run_personne.font.name = "Georgia"
    run_personne.italic = True
    run_personne.bold = True

    # Ajoute les images (assurez-vous que les chemins sont corrects)
    try:
        # p_image = document.add_paragraph()
        # p_image.alignment = 
        document.add_picture(str(PROJECT_ROOT / 'resources' / 'image1.png'), width=Pt(200))
        # Permet d'accéder au dernier paragraphe et de le centrer
        last_paragraph = document.paragraphs[-1] 
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        document.add_paragraph() # Espace entre les images
        document.add_picture(str(PROJECT_ROOT / 'resources'/ 'image2.png'), width=Pt(200))
        last_paragraph = document.paragraphs[-1]
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        # run_image = paragraph.add_run()
        # inline_shapes = run
    except FileNotFoundError:
        document.add_paragraph("(Images non trouvées, veuillez vérifier les chemins)")

    # Ajoute un saut de page pour que le contenu suivant commence sur une nouvelle page
    document.add_page_break()


def generate_docx(report_path: Optional[str] = None, output_path: Optional[str] = None, json_path: Optional[str] = None, indexNoms: Optional[bool] = False,indexLieux: Optional[bool] = False,indexProfessions: Optional[bool] = False) -> Path:
    """
    Orchestre la deuxième étape : la génération du document Word (.docx).
    1. S'assure que le rapport intermédiaire existe (le génère si besoin).
    2. Crée un document Word avec en-tête et pied de page.
    3. Pour chaque personne dans le rapport, ajoute ses informations formatées.
       - Utilise `_dict_to_text_lines` pour obtenir les lignes de texte avec indentation.
    4. Sauvegarde le document final.
    """
    output_file = Path(output_path) if output_path else DEFAULT_OUTPUT

    if report_path is None:
        report = extract_report(json_path=json_path, report_path=str(DEFAULT_REPORT))
        report_file = DEFAULT_REPORT
    else:
        report_file = Path(report_path)
        if json_path:
            report = extract_report(json_path=json_path, report_path=str(report_file))
        else:
            report = _load_json(report_file)

    document = Document()

    # Création des différents styles de Titre
    style = document.styles.add_style('Titre 1', WD_STYLE_TYPE.PARAGRAPH)
    font = style.font
    font.name = 'Garamond'
    font.size = Pt(20)
    font.bold = True

    # font.bold = True
    # if indexB == True:
    #     index = {}        

    first_person_name = "Inconnu"
    if report.get('people'):
        first_person_name = report['people'][0].get('title', 'Inconnu')
    # Ordre d'apparition de chaque personne dans le document (utilisé pour trier
    # les renvois de page multiples dans le bon ordre).
    order_by_id = {
        p.get('id'): p.get('order', 0)
        for p in report.get('people', [])
        if p.get('id')
    }
    _create_cover_page(document, first_person_name)
    # ------------------------------------

    _set_header_footer(document, 'GEDCOM JSON report')

    title_paragraph = document.add_paragraph()
    title_paragraph.add_run(f"\n\n\n\n\n")
    title_run = _add_heading(document, f"Ancêtres de \n{first_person_name}", 1)

    title_run.runs[0].font.size = Pt(44)
    title_run.runs[0].font.name = "Garamond"
    title_paragraph.paragraph_format.alignment =  WD_ALIGN_PARAGRAPH.RIGHT
    add_horizontal_line(title_paragraph)
    p = document.add_paragraph()
    p.add_run("\n")


    generation_paragraph = document.add_paragraph()
    genration_run = generation_paragraph.add_run(f"... générations sur plus de ... ans")
    add_horizontal_line(generation_paragraph)
    genration_run.font.name = "Garamond"
    genration_run.font.size = Pt(19)
    genration_run.italic = True
    generation_paragraph.paragraph_format.alignment =  WD_ALIGN_PARAGRAPH.RIGHT
    p = document.add_paragraph()
    p.add_run("\n")




    # Table des matières : on réserve sa place ici, et on la remplit à la fin
    # (une fois que toutes les sections existent), voir _fill_table_of_contents.
    document.add_page_break()
    toc_title = _add_heading(document, "Table des matières", 1)
    setStyle(toc_title, "Titre", WD_ALIGN_PARAGRAPH.LEFT)
    toc_title.runs[0].bold = False
    toc_anchor = document.add_paragraph()
    toc_entries: List[Tuple[str, str]] = []   # (texte affiché, nom du signet)
    document.add_page_break()

    # Nombre de personnes connues par génération (pour "Cette section présente N ancêtres…")
    people_per_gen: Dict[int, int] = {}
    for person in report.get('people', []):
        gen = _generation_number(person.get('generation'))
        if gen is not None:
            people_per_gen[gen] = people_per_gen.get(gen, 0) + 1

    last_gen = None
    for person in report.get('people', []):

        generation = _generation_number(person.get('generation'))
        if generation is not None and generation != last_gen:
            if last_gen is not None:
                document.add_page_break()
            last_gen = generation
            meta_paragraph = document.add_paragraph()
            meta_run = meta_paragraph.add_run(f"Génération {generation}")
            meta_run.font.size = Pt(20)
            bookmark = f"toc_gen_{generation}_{len(toc_entries)}"
            _add_bookmark(meta_paragraph, bookmark)
            toc_entries.append((f"Génération {generation}", bookmark))

            if generation > 1:
                bound = person.get('bound')
                generation_paragraph = document.add_paragraph()
                periode = f" Ces personnes ont vécu entre les années {bound[0]} et {bound[1]}." if bound else ""
                generation_paragraph.add_run(f"Cette section présente {people_per_gen[generation]} ancêtres de Didier Michel connus à la {generation}ème génération.{periode} ")

        title_paragraph = document.add_paragraph()
        title_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        title_run = title_paragraph.add_run(f"{person.get('order', '')}. {person.get('title', person.get('id', 'Unknown person'))}")
        title_run.bold = True
        title_run.font.size = Pt(14)
        if person.get('id'):
            _add_bookmark(title_paragraph, _bookmark_name(person['id']))



        # Toutes les informations de la personne rassemblées en un seul texte,
        # plutôt qu'éclatées entre plusieurs champs/paragraphes.
        summary_paragraph = document.add_paragraph()
        morceaux = person.get('summary', [])
        noms = person.get('nom', [])
        for position, morceau in enumerate(morceaux):
            if morceau:
                summary_paragraph.add_run(morceau)
            if position < len(noms):
                summary_paragraph.add_run(noms[position]).bold = True

        # Les pièces jointes (images, documents) restent affichées à part,
        # car elles ne peuvent pas être fondues dans le texte.
        file_paths = [
            line[len(_FILE_SENTINEL):]
            for line in _dict_to_text_lines(person.get('lines', {}))
            if line.startswith(_FILE_SENTINEL)
        ]
        for file_path in file_paths:
            ext = Path(file_path).suffix.lower()
            if ext in IMAGE_EXTENSIONS:
                try:
                    document.add_picture(file_path, width=Inches(4))
                    document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                except Exception as img_err:
                    p = document.add_paragraph()
                    p.add_run(f"[Image non chargée : {Path(file_path).name} — {img_err}]").italic = True
            else:
                p = document.add_paragraph()
                p.paragraph_format.left_indent = Pt(28)
                p.add_run(f"📄 {Path(file_path).name}")

        document.add_paragraph('')
    document.add_page_break()


    p = _add_heading(document, "Statistiques", 1)
    _register_toc_entry(p, "Statistiques", toc_entries)
    setStyle(p, "Titre", WD_ALIGN_PARAGRAPH.LEFT)
    p.runs[0].bold = False
    document.add_page_break()

    p = _add_heading(document, "Origine du patronyme paternel", 1)
    _register_toc_entry(p, "Origine du patronyme paternel", toc_entries)
    setStyle(p, "Titre", WD_ALIGN_PARAGRAPH.LEFT)
    p.runs[0].bold = False
    document.add_page_break()

    p = _add_heading(document, "Origine du patronyme maternel", 1)
    _register_toc_entry(p, "Origine du patronyme maternel", toc_entries)
    setStyle(p, "Titre", WD_ALIGN_PARAGRAPH.LEFT)
    p.runs[0].bold = False
    document.add_page_break()
    
    # -----------------------------------------
    #                 Index 
    # -----------------------------------------
    if indexNoms:
        document.add_page_break()
        p = document.add_paragraph()
        text_unit = _add_heading(document, 'Index des noms', 1)
        _register_toc_entry(text_unit, 'Index des noms', toc_entries)
        setStyle(text_unit,"Titre")
        new_section = document.add_section(WD_SECTION.CONTINUOUS) # Permet que le texte soit sur deux colonnes 
        set_columns(new_section, num_columns=2)

        # text_unit.style.element.rPr.rFonts.set(qn("w:asciiTheme"), "Times New Roman")
        # text_unit.runs[0].font.name = "Times New Roman"
        # text_unit.runs[0].font.name 
        p.add_run("\n")
        # Penser à ajouter les références pour pouvoir remettre les numéros associés
        # On aplatit d'abord la liste en (nom, id, objet) — une personne peut avoir
        # plusieurs noms (nom de naissance + nom d'épouse, etc.) — puis on trie sur
        # ce nom individuel, et non sur le seul premier nom de chaque personne.
        # Sinon les noms "secondaires" ne sont pas insérés dans l'ordre alphabétique.
        entries = sorted(
            (
                (nom, id_cle, objet)
                for id_cle, objet in people_by_id.items()
                for nom in objet.nom
            ),
            key=lambda entry: entry[0].upper(),
        )
        dictionnaire_alphabet = {"A":[None,[],None,[]],"B":[None,[],None,[]],"C":[None,[],None,[]],"D":[None,[],None,[]],"E":[None,[],None,[]],"F":[None,[],None,[]],"G":[None,[],None,[]],
                            "H":[None,[],None,[]],"I":[None,[],None,[]],"J":[None,[],None,[]],"K":[None,[],None,[]],"L":[None,[],None,[]],"M":[None,[],None,[]],"N":[None,[],None,[]],"O":[None,[],None,[]],
                            "P":[None,[],None,[]],"Q":[None,[],None,[]],"R":[None,[],None,[]],"S":[None,[],None,[]],"T":[None,[],None,[]],"U":[None,[],None,[]],"V":[None,[],None,[]],"W":[None,[],None,[]],
                            "X":[None,[],None,[]],"Y":[None,[],None,[]],"Z":[None,[],None,[]]}

        i = 0
        q = None
        last_name = ""
        # ---- Index des noms -----
        for nom, id_cle, objet in entries:
            lettre = nom[0].upper()
            if dictionnaire_alphabet[lettre][0] == None:
                letter_paragraph = document.add_paragraph()
                dictionnaire_alphabet[lettre][0] = letter_paragraph
                letter_run = dictionnaire_alphabet[lettre][0].add_run(f"\n  {lettre}")
                # letter_paragraph.add_run("\n")
                letter_run.bold = True
                letter_run.font.size = Pt(20)
            if nom not in dictionnaire_alphabet[lettre][1]:
                dictionnaire_alphabet[lettre][1].append(nom)
                name_paragraphe = document.add_paragraph()
                dictionnaire_alphabet[lettre][2] = name_paragraphe

                name_run = dictionnaire_alphabet[lettre][2].add_run(f"{nom}")
                new_paragraph = document.add_paragraph()
                dictionnaire_alphabet[lettre][3].append(new_paragraph)
                # name_run.bold = True
                contenu = dictionnaire_alphabet[lettre][3][-1]
                contenu.add_run(f"{objet.prenom} {objet.nom[objet.nom.index(nom)]} — p. ")
                if id_cle:
                    _add_page_refs(contenu, {id_cle} | mentions_by_person.get(id_cle, set()), order_by_id)
                contenu.add_run("\n")
            else:
                contenu = dictionnaire_alphabet[lettre][3][dictionnaire_alphabet[lettre][1].index(nom)]
                contenu.add_run(f"{objet.prenom} {objet.nom[objet.nom.index(nom)]} — p. ")
                if id_cle:
                    _add_page_refs(contenu, {id_cle} | mentions_by_person.get(id_cle, set()), order_by_id)
                contenu.add_run("\n")

            contenu.paragraph_format.left_indent = Pt(10)
            contenu.paragraph_format.space_after = Pt(0.0)
            # print(f"ID : {id_cle}")
            i+=1
        new_section = document.add_section(WD_SECTION.CONTINUOUS) # Permet que le texte soit sur deux colonnes 
        set_columns(new_section, num_columns=1)
    # Même principe que pour l'index des noms : on aplatit en (lieu, id, objet)
    # avant de trier, car une personne peut avoir plusieurs lieux associés.
    if indexLieux:
        document.add_page_break()
        indexLieu = _add_heading(document, "Index des lieux", 1)
        _register_toc_entry(indexLieu, "Index des lieux", toc_entries)
        p = document.add_paragraph()
        new_section = document.add_section(WD_SECTION.CONTINUOUS) # Permet que le texte soit sur deux colonnes 
        set_columns(new_section, num_columns=2)
        entries = sorted(
            (
                (enlever_accents(lieu[0]), lieu, id_cle, objet)
                for id_cle, objet in people_by_id.items()
                for lieu in (objet.lieu or [])
            ),
            key=lambda entry: entry[0].upper(),
        )
        dictionnaire_alphabet = {"A":[None,[],None,[]],"B":[None,[],None,[]],"C":[None,[],None,[]],"D":[None,[],None,[]],"E":[None,[],None,[]],"F":[None,[],None,[]],"G":[None,[],None,[]],
                            "H":[None,[],None,[]],"I":[None,[],None,[]],"J":[None,[],None,[]],"K":[None,[],None,[]],"L":[None,[],None,[]],"M":[None,[],None,[]],"N":[None,[],None,[]],"O":[None,[],None,[]],
                            "P":[None,[],None,[]],"Q":[None,[],None,[]],"R":[None,[],None,[]],"S":[None,[],None,[]],"T":[None,[],None,[]],"U":[None,[],None,[]],"V":[None,[],None,[]],"W":[None,[],None,[]],
                            "X":[None,[],None,[]],"Y":[None,[],None,[]],"Z":[None,[],None,[]]}
        for lieu_nom, lieu_tuple, id_cle, objet in entries:
            _, raison, annee = lieu_tuple
            lettre = enlever_accents(lieu_nom[0].upper())
            if ord(lettre) >= 91 or ord(lettre) <= 64:
                print(f"Erreur de chargement de cette lettre : {lettre}")
                lettre = "Z"
            if dictionnaire_alphabet[lettre][0] == None:
                letter_paragraph = document.add_paragraph()
                dictionnaire_alphabet[lettre][0] = letter_paragraph
                letter_run = dictionnaire_alphabet[lettre][0].add_run(f"  {lettre}")
                letter_run.bold = True
                letter_run.font.size = Pt(20)

            adresse_normalisee = normaliser_adresse(lieu_nom)
            nom_personne = objet.nom[0] if objet.nom else ''
            if adresse_normalisee not in dictionnaire_alphabet[lettre][1]:
                dictionnaire_alphabet[lettre][1].append(adresse_normalisee)
                name_paragraphe = document.add_paragraph()
                dictionnaire_alphabet[lettre][2] = name_paragraphe
                name_run = dictionnaire_alphabet[lettre][2].add_run(f"{lieu_nom}")
                # ----- Styles -----
                name_paragraphe.paragraph_format.line_spacing = 1.1
                name_paragraphe.paragraph_format.space_after = Pt(0.0)
                name_run.font.size = Pt(12)
                # ------------------
                new_paragraph = document.add_paragraph()
                dictionnaire_alphabet[lettre][3].append(new_paragraph)


            # Toujours ajouter cette ligne au paragraphe du lieu (nouveau ou déjà connu),
            # pour lister tous les événements survenus à cet endroit, pas seulement le premier.
            contenu = dictionnaire_alphabet[lettre][3][dictionnaire_alphabet[lettre][1].index(adresse_normalisee)]
            contenu.paragraph_format.left_indent = Pt(10)
            contenu.paragraph_format.space_after = Pt(0.0)
            if raison == "Mariage":
                parts = annee.split()
                annee_seule = parts[-1] if parts else ""
                conjoint = ' '.join(parts[:-1])
                contenu.add_run(f"{annee_seule} - Mariage de {objet.prenom} {nom_personne} et {conjoint}")
            elif raison == "Naissance":
                contenu.add_run(f"{annee} - Naissance {_gender_d(objet.prenom)}{objet.prenom} {nom_personne}")
            elif raison == "Décès":
                contenu.add_run(f"{annee} - Décès {_gender_d(objet.prenom)}{objet.prenom} {nom_personne}")
            else:
                contenu.add_run("DNF")
            contenu.add_run(" — p. ")
            if id_cle:
                _add_page_refs(contenu, {id_cle} | mentions_by_person.get(id_cle, set()), order_by_id)
            contenu.add_run("\n")
        
        setStyle(indexLieu, "Titre")

    if indexProfessions:
        document.add_page_break()
        new_section = document.add_section(WD_SECTION.CONTINUOUS) # Permet que le texte soit sur deux colonnes 
        set_columns(new_section, num_columns=1)
        indexProfession = _add_heading(document, "Index des professions", 1)
        _register_toc_entry(indexProfession, "Index des professions", toc_entries)
        setStyle(indexProfession, "Titre")

        p = document.add_paragraph()
        # p.paragraph_format.line_spacing = 0.15
        # p.paragraph_format.space_after = Pt(0.15)
        new_section = document.add_section(WD_SECTION.CONTINUOUS) # Permet que le texte soit sur deux colonnes 
        set_columns(new_section, num_columns=2)
        dictionnaire_alphabet = {"A":[None,[],None,[]],"B":[None,[],None,[]],"C":[None,[],None,[]],"D":[None,[],None,[]],"E":[None,[],None,[]],"F":[None,[],None,[]],"G":[None,[],None,[]],
                            "H":[None,[],None,[]],"I":[None,[],None,[]],"J":[None,[],None,[]],"K":[None,[],None,[]],"L":[None,[],None,[]],"M":[None,[],None,[]],"N":[None,[],None,[]],"O":[None,[],None,[]],
                            "P":[None,[],None,[]],"Q":[None,[],None,[]],"R":[None,[],None,[]],"S":[None,[],None,[]],"T":[None,[],None,[]],"U":[None,[],None,[]],"V":[None,[],None,[]],"W":[None,[],None,[]],
                            "X":[None,[],None,[]],"Y":[None,[],None,[]],"Z":[None,[],None,[]]}
        entries = sorted(
            (
                (objet.profession, id_cle, objet)
                for id_cle, objet in people_by_id.items()
                if objet.profession
            ),
            key=lambda entry: entry[0].upper(),
        )

        for profession, id_cle, objet in entries:
            lettre = enlever_accents(profession[0].upper())
            if ord(lettre) >= 91 or ord(lettre) <= 64:
                continue
            if dictionnaire_alphabet[lettre][0] == None:
                letter_paragraph = document.add_paragraph()
                dictionnaire_alphabet[lettre][0] = letter_paragraph
                letter_run = dictionnaire_alphabet[lettre][0].add_run(f"\n  {lettre}")
                letter_run.bold = True
                letter_run.font.size = Pt(20)

            nom_personne = objet.nom[0] if objet.nom else ''
            if profession not in dictionnaire_alphabet[lettre][1]:
                dictionnaire_alphabet[lettre][1].append(profession)
                name_paragraphe = document.add_paragraph()
                dictionnaire_alphabet[lettre][2] = name_paragraphe
                # ----- Styles -----
                name_paragraphe.paragraph_format.line_spacing = 1.1
                name_paragraphe.paragraph_format.space_after = Pt(0.0)
                name_run.font.size = Pt(12)
                # ------------------
                name_run = dictionnaire_alphabet[lettre][2].add_run(f"{profession.capitalize()}")
                new_paragraph = document.add_paragraph()
                dictionnaire_alphabet[lettre][3].append(new_paragraph)

            # Toujours ajouter cette ligne au paragraphe de la profession (nouvelle ou déjà
            # connue), pour lister toutes les personnes ayant exercé cette profession.
            contenu = dictionnaire_alphabet[lettre][3][dictionnaire_alphabet[lettre][1].index(profession)]
            contenu.paragraph_format.left_indent = Pt(10)
            contenu.paragraph_format.space_after = Pt(0.0)
            contenu.add_run(f"{objet.prenom} {nom_personne}").font.size = Pt(10)

    # -------------------------------------
    # Ajuste les MARGES pour un livre
    # -------------------------------------
    sections = document.sections
    for section in sections:
        section.top_margin = Inches(1.18)
        section.bottom_margin = Inches(0.99)
        section.left_margin = Inches(1.62)
        section.right_margin = Inches(0.79)

    # Remplit la table des matières (maintenant que toutes les sections et les marges existent)
    first_section = document.sections[0]
    _fill_table_of_contents(toc_anchor, toc_entries,
                            first_section.page_width - first_section.left_margin - first_section.right_margin)

    # Force Word à recalculer tous les champs (PAGE, PAGEREF, ...) à l'ouverture,
    # sans quoi les renvois de page restent affichés avec leur valeur de repli
    # tant que l'utilisateur ne fait pas Ctrl+A puis F9.
    settings_el = document.settings.element
    update_fields = OxmlElement('w:updateFields')
    update_fields.set(qn('w:val'), 'true')
    settings_el.append(update_fields)

    document.save(str(output_file))
    print(f"Intermediate report: {report_file}")
    print(f"Document generated: {output_file}")
    return output_file

def _add_heading(document: Document, text: str, level: int = 1) -> Paragraph:
    """Ajoute un titre de niveau `level` (style Word "Heading N", ou "Title" pour 0).

    Remplace `document.add_heading_with_level`, qui n'existait que dans une copie
    modifiée de python-docx : ici on n'utilise que l'API standard, donc ça marche
    avec une installation normale (Mac, Windows, pipeline GitHub).
    """
    if not 0 <= level <= 9:
        raise ValueError(f"level must be in range 0-9, got {level}")
    style = "Title" if level == 0 else f"Heading {level}"
    return document.add_paragraph(text, style)


def _generation_number(generation: Any) -> Optional[int]:
    """La génération est 1 pour la personne de référence, puis un tuple (génération, …) : renvoie juste le numéro."""
    if isinstance(generation, (tuple, list)):
        return generation[0] if generation else None
    return generation


def _register_toc_entry(paragraph: Paragraph, text: str, toc_entries: List[Tuple[str, str]]) -> None:
    """Pose un signet sur ce titre et l'ajoute à la table des matières."""
    bookmark = f"toc_{len(toc_entries)}"
    _add_bookmark(paragraph, bookmark)
    toc_entries.append((text, bookmark))


def _fill_table_of_contents(anchor: Paragraph, toc_entries: List[Tuple[str, str]], text_width) -> None:
    """
    Écrit la table des matières juste avant `anchor` : une ligne par section,
    "Titre ........ page", le titre étant un lien cliquable vers la section.
    Les numéros de page sont des champs PAGEREF, calculés par Word à l'ouverture
    (comme pour les index).
    """
    for text, bookmark in toc_entries:
        paragraph = anchor.insert_paragraph_before()
        paragraph.paragraph_format.tab_stops.add_tab_stop(
            text_width, WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        paragraph.paragraph_format.space_after = Pt(4)

        # Titre cliquable : le run est placé dans un <w:hyperlink w:anchor="signet">
        run = paragraph.add_run(text)
        hyperlink = OxmlElement('w:hyperlink')
        hyperlink.set(qn('w:anchor'), bookmark)
        hyperlink.set(qn('w:history'), '1')
        run._r.addprevious(hyperlink)
        hyperlink.append(run._r)

        paragraph.add_run('	')
        _add_page_ref(paragraph, bookmark)

        # paragraph.runs ne voit pas le run déplacé dans le lien : on l'ajoute à la main
        for r in [run] + list(paragraph.runs):
            r.font.name = "Garamond"
            r.font.size = Pt(13)


def setStyle(paragraph, style: str, alignment=WD_ALIGN_PARAGRAPH.LEFT):
    paragraph.runs[0].font.name = "Garamond"
    paragraph.runs[0].font.color.rgb = RGBColor(0,0,0)
    if style == "Titre":
        if alignment != None:
            paragraph.paragraph_format.alignment = alignment
        else:
            paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.runs[0].font.size = Pt(20)
        paragraph.runs[0].bold = True
        add_horizontal_line(paragraph)
    if style == "Titre 1":
        paragraph.runs[0].font.size = Pt(16)
        paragraph.runs[0].bold = True
    elif style == "Titre 2":
        paragraph.runs[0].font.size = Pt(16)
    elif style == "Titre 3":
        paragraph.runs[0].font.size = Pt(14)
        
def add_horizontal_line(paragraph):
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    pPr.append(pBdr)

    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')      # épaisseur (en 1/8 pt)
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '000000')  # couleur (noir ici)
    pBdr.append(bottom)

    return paragraph


def enlever_accents(texte):
    """Enlève l'accent de la première lettre de `texte` uniquement, le reste est inchangé."""
    if not texte:
        return texte
    nfkd = unicodedata.normalize('NFKD', texte[0])
    premiere_lettre = ''.join(c for c in nfkd if not unicodedata.combining(c))
    return premiere_lettre + texte[1:]


def normaliser_adresse(adresse):
    """Normalise une adresse pour permettre la comparaison entre deux adresses
    (espaces retirés, virgule finale retirée). Ne modifie pas l'adresse d'origine :
    à utiliser uniquement pour comparer, jamais pour l'affichage."""
    if not adresse:
        return adresse
    normalisee = adresse.replace(", ", ",")
    normalisee = adresse.replace(" ,", ",")
    normalisee = normalisee.rstrip(",")
    return normalisee

def set_columns(section, num_columns=2, space=720, equal_width=True):
    sectPr = section._sectPr
    cols = sectPr.find(qn('w:cols'))
    if cols is None:
        cols = OxmlElement('w:cols')
        sectPr.append(cols)
    cols.set(qn('w:num'), str(num_columns))
    cols.set(qn('w:space'), str(space))  # espace entre colonnes, en twips (720 = 0,5 cm... en fait 1/1440 pouce, donc 720 = 0.5 inch)
    if equal_width:
        cols.set(qn('w:equalWidth'), '1')

if __name__ == '__main__':
    # --- Point d'entrée du script ---
    # Cette section permet d'exécuter le script depuis la ligne de commande
    # en passant des arguments pour les chemins des fichiers.
    import argparse

    parser = argparse.ArgumentParser(description='Extract people from a JSON file and generate a DOCX report')
    parser.add_argument('--json', help='Path to the input JSON file', default=str(DEFAULT_JSON))
    parser.add_argument('--report', help='Path to the intermediate extracted report JSON', default=str(DEFAULT_REPORT))
    parser.add_argument('--out', help='Output DOCX path', default=str(DEFAULT_OUTPUT))
    args = parser.parse_args()
    generate_docx(report_path=args.report, output_path=args.out, json_path=args.json)
