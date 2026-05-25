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
from pathlib import Path
from typing import Any, Dict, List, Optional
from datetime import date

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

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
        return _display_name(target)
    if target_type == 'bio:Relationship':
        return _family_label(str(reference_id), index)

    title = target.get('TITL')
    if title:
        return str(title)

    name_value = target.get('foaf:name')
    if isinstance(name_value, str):
        return name_value

    if isinstance(name_value, dict):
        return _display_name(target)

    return str(reference_id)


def _resolve_reference(reference_id: str, index: Dict[str, Dict[str, Any]], current_node: Dict[str, Any]) -> Dict[str, Any]:
    """
    Résout une référence ID en un objet structuré contenant des informations clés.
    Va chercher l'objet correspondant à l'ID dans l'index et retourne un dictionnaire
    avec son nom, son type (personne, famille), etc.
    """
    target = index.get(str(reference_id), {})
    resolved: Dict[str, Any] = {'@id': reference_id, 'display': _reference_summary(reference_id, index)}

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
        resolved['participants'] = [
            _display_name(index[p.get('@id')])
            for p in participants
            if isinstance(p, dict) and p.get('@id') in index
        ]

    elif target.get('@type') =='_:SOUR':
        if target.get('@type') == '_:SOUR':
            # resolved['title'] = title
            resolved['type_source'] = target.get('TYPE')
            resolved['type_source'] = target.get('WWW')
            if target.get('DATA'):
                t = target.get('DATA')
                u = t.get('TEXT')
                u = u.values
                for e in u:
                    resolved['texte'] += e
        print("je suis ici")
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
    place_info = event.get('bio:place')
    place_value = ''
    if isinstance(place_info, dict):
        # Si le lieu est un objet, on cherche la valeur dans 'GEDVALUE'
        place_value = str(place_info.get('GEDVALUE', '')).strip()
    elif place_info:
        # Sinon, on utilise la valeur directement
        place_value = str(place_info).strip()

    # Construire la phrase en plusieurs parties
    parts = [person_name, person_gender_prefix]
    if date_value:
        parts.append(f"le {date_value}")
    if place_value:
        parts.append(f"à {place_value}")

    # Joindre les parties. 'est né' ou 'est née'
    if len(parts) > 2:
        return " ".join(parts)
    
    # S'il n'y a ni date ni lieu, on retourne une chaîne vide ou une info de base.
    return ""


def _pretty_nested_value(key: str, value: Any, index: Dict[str, Dict[str, Any]], current_node: Dict[str, Any]) -> Any:
    """
    Fonction récursive au cœur de la transformation des données.
    Parcourt chaque champ d'un objet et le transforme en une valeur "jolie" et lisible.
    - Remplace les références ID par des résumés.
    - Formate les événements (naissance, mariage) en phrases descriptives.
    - Ajoute des labels 'display' pour un affichage direct.
    """
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
        role = 'enfant de' if key == 'FAMC' else _gender_role(current_node)
        display = f"{_display_name(current_node)} est {role} la famille {family_name}"

        result: Dict[str, Any] = {'display': display}
        if family_id is not None:
            result['ofFamily'] = _resolve_reference(str(family_id), index, current_node)
        if isinstance(value, dict):
            for sub_key, sub_value in value.items():
                if sub_key != 'ofFamily':
                    result[sub_key] = _pretty_nested_value(sub_key, sub_value, index, current_node)
        return result

    if key == 'bio:event':
        events = value if isinstance(value, list) else [value]
        pretty_events: List[Any] = []
        for event in events:
            if not isinstance(event, dict):
                pretty_events.append(event)
                continue

            event_type = str(event.get('@type', ''))
            date_value = str(event.get('dc:date', '')).strip()
            display = None
            if event_type == 'bio:Birth':
                display = _format_birth_event(event, _display_name(current_node), _gender_prefix(current_node))
            elif event_type == 'bio:Death' and date_value:
                display = f"{_display_name(current_node)} est décédé(e) le {date_value}"
            elif event_type == 'bio:Marriage' and date_value:
                display = f"{_display_name(current_node)} s'est marié(e) le {date_value}"
            elif date_value:
                display = f"Événement daté du {date_value}"

            pretty_event: Dict[str, Any] = {}
            if display:
                pretty_event['display'] = display
            for sub_key, sub_value in event.items():
                # On ignore les clés 'display' (que l'on vient de créer) et 'MAP'
                if sub_key not in {'display', 'MAP'}:
                    pretty_event[sub_key] = _pretty_nested_value(sub_key, sub_value, index, current_node)
            pretty_events.append(pretty_event)

        return pretty_events if isinstance(value, list) else pretty_events[0]

    if isinstance(value, str) and ID_PATTERN.match(value):
        return _resolve_reference(value, index, current_node)

    if isinstance(value, dict) and '@id' in value and len(value) == 1:
        reference_id = str(value.get('@id'))
        return _resolve_reference(reference_id, index, current_node)

    if key == 'OBJE':
        entries = value if isinstance(value, list) else [value]
        pretty_media: List[Any] = []
        for entry in entries:
            if not isinstance(entry, dict):
                pretty_media.append(entry)
                continue

            media_item: Dict[str, Any] = {}
            media_id = entry.get('@id')
            if media_id:
                media_item['@id'] = media_id
                target = index.get(str(media_id), {})
                title = target.get('TITL')
                file_value = None
                file_info = target.get('FILE')
                if isinstance(file_info, dict):
                    file_value = file_info.get('GEDVALUE')
                display_bits = [f"Pièce jointe {media_id}"]
                if title:
                    display_bits.append(str(title))
                if file_value:
                    display_bits.append(str(file_value))
                media_item['display'] = ' - '.join(display_bits)
            if entry.get('_PRIM'):
                media_item['_PRIM'] = entry.get('_PRIM')
            pretty_media.append(media_item)
        return pretty_media if isinstance(value, list) else pretty_media[0]

    if isinstance(value, dict):
        result: Dict[str, Any] = {}
        for sub_key, sub_value in value.items():
            result[sub_key] = _pretty_nested_value(sub_key, sub_value, index, current_node)
        return result

    if isinstance(value, list):
        if all(isinstance(item, (str, int, float, bool, type(None))) for item in value):
            return value if len(value) > 1 else (value[0] if value else '')

        return [_pretty_nested_value(key, item, index, current_node) for item in value]

    return value


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
            print(f"Traitement de la personne {order}/{len(graph)} : {node_id} - {_display_name(node)}")

            heading = _display_name(node)

            person_dict: Dict[str, Any] = {}
            for key, value in node.items():
                if key in {'@id', '@type'}:
                    continue
                person_dict[key] = _pretty_nested_value(key, value, index, node)

            report['people'].append(
                {
                    'order': order,
                    'id': node.get('@id'),
                    'type': node.get('@type'),
                    'title': str(heading),
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


def _dict_to_text_lines(data: Any, indent: int = 0) -> List[str]:
    """Recursively convert a dict/value structure back to indented text lines for display."""
    prefix = '  ' * indent
    lines: List[str] = []
    
    if isinstance(data, dict):
        if 'display' in data:
            lines.append(f"{prefix}{data['display']}")
        for key, value in data.items():
            if key == 'display':
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
    run_titre.font.size = Pt(36)
    run_titre.bold = True

    # Ajoute le nom de la première personne
    p_personne = document.add_paragraph()
    p_personne.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_personne = p_personne.add_run(f"de {first_person_name}\n\n")
    run_personne.font.size = Pt(18)

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


def generate_docx(report_path: Optional[str] = None, output_path: Optional[str] = None, json_path: Optional[str] = None) -> Path:
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
        meta_run = meta_paragraph.add_run(f"@id: {person.get('id', '')}    @type: {person.get('type', '')}")
        meta_run.italic = True
        meta_run.font.size = Pt(9)

        # Convert dict to text lines for display
        person_lines = _dict_to_text_lines(person.get('lines', {}))
        
        for line in person_lines:
            indent = 0
            content = line
            while content.startswith('  '):
                indent += 1
                content = content[2:]

            paragraph = document.add_paragraph()
            paragraph.paragraph_format.left_indent = Pt(14 * indent)
            paragraph.add_run(content)

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
