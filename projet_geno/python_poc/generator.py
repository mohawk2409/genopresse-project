import json
import re
from pathlib import Path
from typing import Any, Dict, List, Optional

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_JSON = PROJECT_ROOT / 'python_poc' / 'test.json'
DEFAULT_REPORT = PROJECT_ROOT / 'python_poc' / 'extracted_report.json'
DEFAULT_OUTPUT = PROJECT_ROOT / 'python_poc' / 'output.docx'
ID_PATTERN = re.compile(r'^_:[A-Za-z0-9]+$')


def _add_page_number(paragraph):
    paragraph.add_run('Page ')
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), 'PAGE \\* MERGEFORMAT')
    paragraph._p.append(fld)
    paragraph.add_run(' of ')
    fld2 = OxmlElement('w:fldSimple')
    fld2.set(qn('w:instr'), 'NUMPAGES \\* MERGEFORMAT')
    paragraph._p.append(fld2)


def _load_json(path: Path) -> Dict[str, Any]:
    for encoding in ['utf-8-sig', 'utf-16', 'utf-8', 'latin-1']:
        try:
            with open(path, 'r', encoding=encoding) as file_handle:
                return json.load(file_handle)
        except (UnicodeDecodeError, json.JSONDecodeError):
            continue
    raise ValueError(f"Could not decode {path} with any encoding")


def _save_json(path: Path, payload: Dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, 'w', encoding='utf-8') as file_handle:
        json.dump(payload, file_handle, ensure_ascii=False, indent=2)


def _node_index(root: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    index: Dict[str, Dict[str, Any]] = {}
    for node in root.get('@graph', []):
        if isinstance(node, dict) and '@id' in node:
            index[str(node['@id'])] = node
    return index


def _display_name(node: Dict[str, Any]) -> str:
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
    gender = str(node.get('foaf:gender', '')).upper()
    if gender == 'F':
        return 'née'
    if gender == 'M':
        return 'né'
    return 'né(e)'


def _gender_role(node: Dict[str, Any]) -> str:
    gender = str(node.get('foaf:gender', '')).upper()
    if gender == 'F':
        return 'l\'épouse'
    if gender == 'M':
        return 'l\'époux'
    return 'le conjoint'


def _family_label(family_id: str, index: Dict[str, Dict[str, Any]]) -> str:
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
    else:
        title = target.get('TITL')
        if title:
            resolved['title'] = title

    if current_node.get('@id') == reference_id:
        resolved['self'] = True

    return resolved


def _pretty_nested_value(key: str, value: Any, index: Dict[str, Dict[str, Any]], current_node: Dict[str, Any]) -> Any:
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
            if event_type == 'bio:Birth' and date_value:
                display = f"{_display_name(current_node)} est {_gender_prefix(current_node)} le {date_value}"
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
                if sub_key != 'display':
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
    json_file = Path(json_path) if json_path else DEFAULT_JSON
    report_file = Path(report_path) if report_path else DEFAULT_REPORT

    root = _load_json(json_file)
    index = _node_index(root)

    report: Dict[str, Any] = {
        'sourceFile': str(json_file),
        'people': []
    }

    for order, node in enumerate(root.get('@graph', []), start=1):
        if not isinstance(node, dict) or node.get('@type') != 'foaf:Person':
            continue

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

    _save_json(report_file, report)
    return report


def _set_header_footer(document: Document, title: str) -> None:
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


def generate_docx(report_path: Optional[str] = None, output_path: Optional[str] = None, json_path: Optional[str] = None) -> Path:
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
    import argparse

    parser = argparse.ArgumentParser(description='Extract people from a JSON file and generate a DOCX report')
    parser.add_argument('--json', help='Path to the input JSON file', default=str(DEFAULT_JSON))
    parser.add_argument('--report', help='Path to the intermediate extracted report JSON', default=str(DEFAULT_REPORT))
    parser.add_argument('--out', help='Output DOCX path', default=str(DEFAULT_OUTPUT))
    args = parser.parse_args()

    generate_docx(report_path=args.report, output_path=args.out, json_path=args.json)
