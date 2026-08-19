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
from typing import Any, Dict, List, Optional, Tuple
from datetime import date

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

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

def _display_name(node: Dict[str, Any]) -> str:
    """
    Extrait et formate un nom complet et lisible pour une personne (nœud).
    Combine le prénom (GIVN) et le nom (SURN) ou utilise d'autres champs si indisponibles.
    """
    name_value = node.get('foaf:name')
    if isinstance(name_value, dict):
        given = str(name_value.get('GIVN') or '').strip()
        surname = str(name_value.get('SURN') or '').strip()
        gedvalue = str(name_value.get('GEDVALUE') or '').strip()
        if given and surname:
            return f"{given} {surname}"
        if gedvalue:
            return gedvalue.replace('/', ' ').replace('  ', ' ').strip()
    if isinstance(name_value, str) and name_value.strip():
        return name_value.strip().replace('/', ' ').replace('  ', ' ')
    return str(node.get('@id', 'Person'))


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

def _gender_d(node: Dict[str, Any]) -> str:
    """Retourne un rôle accordé en genre ("l'époux", "l'épouse")."""
    gender = node.get('foaf:name', '')
    prefixe = str(gender.get('SURN', '')).upper()
    if prefixe[0] == 'A' or prefixe[0] == 'E' or prefixe[0] == 'I' or prefixe[0] == 'O' or prefixe[0] == 'U' or prefixe[0] == 'Y':
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



def _family_label(family_id: str, index: Dict[str, Dict[str, Any]]) -> str:
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
                names.append(_display_name(index[participant_id]))

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


def _reference_summary(reference_id: str, index: Dict[str, Dict[str, Any]]) -> str:
    """
    Génère un résumé textuel pour un objet référencé par son ID.
    Par exemple, pour l'ID d'une personne, retourne son nom ; pour une famille, son label.
    """
    target = index.get(str(reference_id), {})
    if not target:
        return str(reference_id)

    target_type = str(target.get('@type', ''))
    if target_type == 'foaf:Person':
        final_texte = _display_name(target)
        marriage_phrase = _marriage_phrase(target, index)
        if marriage_phrase:
            final_texte += f", {marriage_phrase}"
        return final_texte
    if target_type == 'bio:Relationship':
        return _family_label(str(reference_id), index)
        # if target_type == 'bio:IndividualEvent':
        #     final_texte += _note_label(str(reference_id), index)
        # return final_texte


    title = target.get('TITL') or target.get('ABBR')
    if title:
        return str(title)

    name_value = target.get('foaf:name')
    if isinstance(name_value, str):
        return name_value

    if isinstance(name_value, dict):
        return _display_name(target)

    return str(reference_id)


def _person_marriages_in_order(person: Dict[str, Any], index: Dict[str, Dict[str, Any]]) -> List[str]:
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
    family_ids = _person_marriages_in_order(person, index)
    if not family_ids:
        return ''

    family = index.get(family_ids[0], {})
    participants = family.get('bio:participant', [])
    if isinstance(participants, dict):
        participants = [participants]
    person_id = person.get('@id')
    spouse_id = next(
        (p.get('@id') for p in participants
         if isinstance(p, dict) and p.get('@id') != person_id and p.get('@id') in index),
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
        place_value = _event_place(marriage_event)
        bits = [b for b in [date_phrase, f"à {place_value}" if place_value else ''] if b]
        marriage_phrase = ' '.join(bits)

    spouse_name = _display_name(index[spouse_id])
    verb_phrase = _gender_marriage(person)
    if marriage_phrase:
        return f"{verb_phrase} avec {spouse_name} {marriage_phrase}"
    return f"{verb_phrase} avec {spouse_name}"


def _resolve_reference(reference_id: str, index: Dict[str, Dict[str, Any]], current_node: Dict[str, Any]) -> Dict[str, Any]:
    """
    Résout une référence ID en un objet structuré contenant des informations clés.
    Va chercher l'objet correspondant à l'ID dans l'index et retourne un dictionnaire
    avec son nom, son type (personne, famille), etc.
    """
    target = index.get(str(reference_id), {})
    resolved: Dict[str, Any] = {'@id': reference_id, 'display': _reference_summary(reference_id, index)}
    # work_value = _display_work(current_node)

    if target.get('@type') == 'foaf:Person':
        resolved['type'] = 'person'
        resolved['name'] = _display_name(target)
        gender = target.get('foaf:gender')
        if gender:
            resolved['gender'] = gender
    elif target.get('@type') == 'bio:Relationship':
        resolved['type'] = 'family'
        resolved['name'] = _family_label(str(reference_id), index)
        participants = target.get('bio:participant', [])
        if isinstance(participants, dict):
            participants = [participants]
        participant_ids = [
            p.get('@id') for p in participants
            if isinstance(p, dict) and p.get('@id') in index
        ]
        resolved['participants'] = [_display_name(index[pid]) for pid in participant_ids]

        # Si la personne courante est l'un des conjoints de cette famille, on construit
        # une description de cette union précise : phrase de mariage (ou remariage selon
        # son rang chronologique parmi tous ses mariages), biographie du conjoint et
        # éventuels enfants de cette union. Une seule famille est décrite ici : celle
        # référencée par reference_id, pour éviter de dupliquer l'historique marital
        # complet dans chaque entrée de 'bio:relationship'.
        current_id = current_node.get('@id')
        if current_id in participant_ids:
            description_parts = []

            ordered_family_ids = _person_marriages_in_order(current_node, index)
            family_id = str(reference_id)
            rank = ordered_family_ids.index(family_id) if family_id in ordered_family_ids else 0

            spouse_id = next((pid for pid in participant_ids if pid != current_id), None)
            if spouse_id:
                marriage_event = target.get('bio:event')
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
                    birth_event = _find_event(current_node, 'bio:Birth')
                    age = _compute_age(
                        str(birth_event.get('dc:date', '')).strip() if birth_event else '',
                        str(marriage_event.get('dc:date', '')).strip(),
                    )

                spouse_name = _display_name(index[spouse_id])
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
        return _format_place(str(place_info.get('GEDVALUE', '')).strip())
    if place_info:
        return _format_place(str(place_info).strip())
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
        if date_phrase:
            death_bits.append(date_phrase)
        place_value = _event_place(death_event)
        if place_value:
            death_bits.append(f"à {place_value}")
        death_sentence = ' '.join(death_bits)
        if age is not None:
            death_sentence += f", à l'âge de {age} ans"
        sentence += f" {death_sentence}."
        
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
    family_ids = _person_marriages_in_order(node, index)

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
    # Skip unwanted metadata keys globally
    if key in SKIP_KEYS:
        return None

    if key == 'foaf:name' and isinstance(value, dict):
        display_name = _display_name(current_node)
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

        family_name = _family_label(str(family_id), index) if family_id else 'inconnue'
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
                print("bio:place", place_info)
                place_value = ''
                if isinstance(place_info, dict):
                    # Si le lieu est un objet, on cherche la valeur dans 'GEDVALUE'
                    place_value = _format_place(str(place_info.get('GEDVALUE', '')).strip())
                elif place_info:
                    # Sinon, on utilise la valeur directement
                    place_value = _format_place(str(place_info).strip())
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
        
        display = f"{_display_name(current_node)}, {role1} {_gender_d(current_node)}{family_name}. {pronoun} est {prefix} le {date_info} à {place_value}."

        # display = f"{_format_birth_event(event, _display_name(current_node), _gender_prefix(current_node))}"

        result: Dict[str, Any] = {'display': display}
        if family_id is not None:
            result['ofFamily'] = _resolve_reference(str(family_id), index, current_node)['display']
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
                place_value = _format_place(str(place_info.get('GEDVALUE', '')).strip())
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
            print(pretty_event)

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
            return _resolve_reference(ref_id, index, current_node)
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

    # 1. Filiation — inchangé
    famc = person_dict.get('FAMC')
    if isinstance(famc, dict) and famc.get('display'):
        filiation = str(famc['display']).split('.', 1)[0].strip()
        if filiation:
            parts.append(filiation + '.')

    events = person_dict.get('bio:event', [])
    if isinstance(events, dict):
        events = [events]
    events = [e for e in events if isinstance(e, dict) and e.get('display')]

    birth_display = next((e['display'] for e in events if e.get('eventType') == 'bio:Birth'), None)
    death_display = next((e['display'] for e in events if e.get('eventType') == 'bio:Death'), None)
    other_displays = [e['display'] for e in events if e.get('eventType') not in {'bio:Birth', 'bio:Death'}]

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
        parts.append(str(death_display).strip())

    # 7. Fratrie
    siblings = person_dict.get('rel:siblingOf')
    if siblings:
        siblings_list = siblings if isinstance(siblings, list) else [siblings]
        names = [
            s.get('display') or s.get('name') for s in siblings_list
            if isinstance(s, dict) and (s.get('display') or s.get('name'))
        ]
        if names:
            intro = "Frère ou sœur :" if len(names) == 1 else "Frères et sœurs :"
            parts.append(f"{intro} {', '.join(str(n) for n in names)}.")

    return ' '.join(p for p in parts if p)


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

    generation_map = _compute_generations(root, index)

    report: Dict[str, Any] = {
        'sourceFile': str(json_file),
        'people': []
    }

    print("Début du traitement du fichier JSON...")
    graph = root.get('@graph', [])
    print(f"Nombre total d'éléments dans '@graph': {len(graph)}")

    for order, node in enumerate(graph, start=1):
        try:
            if not isinstance(node, dict) or node.get('@type') != 'foaf:Person':
                continue

            node_id = node.get('@id', 'ID inconnu')
            generation = generation_map.get(str(node_id))

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

            report['people'].append(
                {
                    'order': order,
                    'id': node.get('@id'),
                    'type': node.get('@type'),
                    'title': str(heading),
                    'generation': generation,
                    'summary': summary,
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

def _compute_generations(root: Dict[str, Any], index: Dict[str, Dict[str, Any]]) -> Dict[str, int]:
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
        return {}

    generations: Dict[str, int] = {ref_id: 1}
    queue = [(ref_id, 1)]

    while queue:
        person_id, gen = queue.pop(0)
        person = index.get(person_id, {})

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
                    generations[parent_id] = gen + 1
                    queue.append((parent_id, gen + 1))

    return generations


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
        document.add_picture(str(PROJECT_ROOT / 'image1.png'), width=Pt(200))
        # Permet d'accéder au dernier paragraphe et de le centrer
        last_paragraph = document.paragraphs[-1] 
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        document.add_paragraph() # Espace entre les images
        document.add_picture(str(PROJECT_ROOT / 'image2.png'), width=Pt(200))
        last_paragraph = document.paragraphs[-1] 
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        # run_image = paragraph.add_run()
        # inline_shapes = run
    except FileNotFoundError:
        document.add_paragraph("(Images non trouvées, veuillez vérifier les chemins)")

    # Ajoute un saut de page pour que le contenu suivant commence sur une nouvelle page
    document.add_page_break()


def generate_docx(report_path: Optional[str] = None, output_path: Optional[str] = None, json_path: Optional[str] = None, index: Optional[bool] = None) -> Path:
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

    # --- Création de la page de garde ---
    
    # if index == True:
        # do smthg

    first_person_name = "Inconnu"
    if report.get('people'):
        first_person_name = report['people'][0].get('title', 'Inconnu')
    _create_cover_page(document, first_person_name)
    # ------------------------------------

    _set_header_footer(document, 'GEDCOM JSON report')

    for person in report.get('people', []):
        title_paragraph = document.add_paragraph()
        title_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        title_run = title_paragraph.add_run(f"{person.get('order', '')}. {person.get('title', person.get('id', 'Unknown person'))}")
        title_run.bold = True
        title_run.font.size = Pt(14)

        meta_paragraph = document.add_paragraph()
        generation = person.get('generation')
        generation_str = f"    Génération: {generation}" if generation is not None else "    Génération: non reliée"
        meta_run = meta_paragraph.add_run(f"@id: {person.get('id', '')}    @type: {person.get('type', '')}{generation_str}")
        meta_run.italic = True
        meta_run.font.size = Pt(9)

        # Toutes les informations de la personne rassemblées en un seul texte,
        # plutôt qu'éclatées entre plusieurs champs/paragraphes.
        summary_paragraph = document.add_paragraph()
        summary_paragraph.add_run(person.get('summary', ''))

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

    document.save(str(output_file))
    print(f"Intermediate report: {report_file}")
    print(f"Document generated: {output_file}")
    return output_file

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
