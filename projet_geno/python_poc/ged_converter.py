"""
Convertisseur GEDCOM (.ged) -> JSON-LD (.json).

Portage Python de converterGedJson (convert.js, ged2Json.js, fileLines.js).
Produit exactement le même JSON que la version Node.js, sans aucune dépendance
externe (uniquement la bibliothèque standard), donc fonctionne sous Windows et macOS.

Utilisation depuis Python :
    from ged_converter import ged_to_json, convert_file
    data = ged_to_json("fichier.ged")             # -> dict
    convert_file("fichier.ged", "fichier.json")   # écrit le fichier

Utilisation en ligne de commande :
    python ged_converter.py fichier.ged [sortie.json]
    (sans fichier de sortie, le JSON est affiché dans le terminal)
"""

import json
import re
import sys
from pathlib import Path
from typing import Any, Dict, Iterator, List, Optional

_MISSING = object()  # équivalent de `undefined` en JavaScript


# ---------------------------------------------------------------------------
# Étape 1 : lecture du fichier ligne par ligne (fileLines.js)
# ---------------------------------------------------------------------------

def _read_lines(path) -> Iterator[str]:
    """Découpe le fichier sur chaque \\n ou \\r (comme fileLines.js)."""
    raw = Path(path).read_bytes()
    text = raw.decode("utf-8", errors="replace")
    yield from re.split(r"[\r\n]", text)


def _js_parse_int(s: str) -> Optional[int]:
    """Équivalent de parseInt() en JS : ignore les espaces/BOM au début, lit les chiffres."""
    m = re.match(r"[+-]?\d+", s.lstrip().lstrip("﻿").lstrip())
    return int(m.group()) if m else None


# ---------------------------------------------------------------------------
# Étape 2 : lignes GEDCOM -> objets JSON bruts (ged2Json.js)
# ---------------------------------------------------------------------------

def _js_gedvalue(item: Any) -> str:
    """Lit item['GEDVALUE'] comme le ferait JS (renvoie 'undefined' si absent)."""
    if isinstance(item, dict) and "GEDVALUE" in item:
        return item["GEDVALUE"]
    return "undefined"


def _ged_simplify(obj: Dict[str, Any]) -> Any:
    keys = list(obj.keys())
    if len(keys) == 1 and keys[0] == "GEDVALUE":
        return obj["GEDVALUE"]

    out: Dict[str, Any] = {}
    for key in keys:
        if not isinstance(obj[key], list):
            obj[key] = [obj[key]]

        if key == "ADDR":
            values = []
            for el in obj[key]:
                if "CONT" in el:
                    el["GEDVALUE"] += "\n" + _js_gedvalue(el["CONT"])
                    del el["CONT"]
                values.append(_ged_simplify(el))
        elif key == "DATE":
            values = []
            for el in obj[key]:
                if "TIME" in el:
                    # ajoute simplement à GEDVALUE la valeur trouvée dans TIME
                    el["GEDVALUE"] += " " + _js_gedvalue(el["TIME"])
                    del el["TIME"]
                values.append(_ged_simplify(el))
        else:
            values = [el if isinstance(el, str) else _ged_simplify(el) for el in obj[key]]

        out[key] = values[0] if len(values) == 1 else values
        if out[key] == "":
            del out[key]
    return out


def _ged_records(lines) -> Iterator[Any]:
    """Regroupe les lignes en un objet par enregistrement de niveau 0."""
    current: Optional[Dict[str, Any]] = None
    last_levels: Dict[int, Dict[str, Any]] = {}

    for line in lines:
        if not line or line.strip() == "":
            continue
        s = line.find(" ")
        level = _js_parse_int(line[0:s])
        s2 = line.find(" ", s + 1)
        if s2 != -1:
            code = line[s + 1:s2]
            value = line[s2 + 1:]
        else:
            code = line[s + 1:]
            value = ""

        if level == 0:
            if current is not None:
                yield _ged_simplify(current)
            current = {"GEDVALUE": line[s + 1:]}
            last_levels = {0: current}
        else:
            parent = last_levels.get(level - 1) if level is not None else None
            if parent is None:
                raise ValueError(f"Ligne GEDCOM invalide (pas de parent pour le niveau {level}) : {line!r}")
            new_item = {"GEDVALUE": value}
            if code not in parent:
                parent[code] = new_item
            else:
                if not isinstance(parent[code], list):
                    parent[code] = [parent[code]]
                parent[code].append(new_item)
            last_levels[level] = new_item

    if current is not None:
        yield _ged_simplify(current)


# ---------------------------------------------------------------------------
# Étape 3 : objets bruts -> JSON-LD (convert.js)
# ---------------------------------------------------------------------------

def _set_property(obj: dict, prop: str, value: Any) -> None:
    """Affecte value ; une liste d'un seul élément devient cet élément, une liste vide supprime."""
    if isinstance(value, list) and len(value) == 0:
        obj.pop(prop, None)
    elif isinstance(value, list) and len(value) == 1:
        obj[prop] = value[0]
    else:
        obj[prop] = value


def _as_array(obj: dict, prop: str) -> list:
    if prop not in obj:
        return []
    return obj[prop] if isinstance(obj[prop], list) else [obj[prop]]


def _set_or_add(obj: dict, prop: str, value: Any) -> None:
    """Affecte la valeur, ou l'ajoute à la valeur existante (qui devient une liste)."""
    if prop not in obj:
        _set_property(obj, prop, value)
    else:
        extra = value if isinstance(value, list) else [value]
        _set_property(obj, prop, _as_array(obj, prop) + extra)


def _parse_property(obj: dict, prop: str, callback) -> None:
    """Applique callback à chaque valeur de la propriété."""
    if prop not in obj:
        return
    if isinstance(obj[prop], list):
        obj[prop] = [callback(el) for el in obj[prop]]
    else:
        obj[prop] = callback(obj[prop])


def _filter_property(obj: dict, prop: str, callback) -> None:
    if prop not in obj:
        return
    _set_property(obj, prop, [el for el in _as_array(obj, prop) if callback(el)])


def _rename_property(obj: dict, old_name: str, new_name: str) -> None:
    """Renomme une propriété ; fusionne si la nouvelle existe déjà."""
    if old_name not in obj:
        return
    _set_or_add(obj, new_name, obj[old_name])
    del obj[old_name]


def _transform_id(raw: str) -> Dict[str, str]:
    """'@I001@ INDI' -> {'@id': '_:I001', '@type': '_:INDI'}"""
    out = {"@id": "", "@type": ""}
    if raw[0] != "@":
        out["@id"] = "_:" + raw.strip()
        return out
    pieces = raw.split(" ")
    out["@id"] = "_:" + pieces[0].replace("@", "")
    if len(pieces) > 1:
        out["@type"] = "_:" + " ".join(pieces[1:])
    return out


def _ld_simplify(obj: dict) -> dict:
    """Convertit un objet GEDCOM brut en objet structuré JSON-LD."""
    out: Dict[str, Any] = {}
    for key in list(obj.keys()):
        if not isinstance(obj[key], list):
            obj[key] = [obj[key]]

        values = []
        for el in obj[key]:
            if isinstance(el, str):
                if el == "":
                    continue
                if key == "GEDVALUE" and el == "HEAD":
                    values.append({"@id": "_:HEAD"})  # HEAD est un cas particulier
                elif el[0] == "@":
                    parsed = _transform_id(el)
                    # GEDVALUE = identifiant de l'objet courant ; sinon c'est une référence
                    values.append(parsed if key == "GEDVALUE" else {"@id": parsed["@id"]})
                else:
                    values.append(el)
            else:
                values.append(_ld_simplify(el))

        out[key] = values[0] if len(values) == 1 else values

        nested = out.get("GEDVALUE", _MISSING)
        if isinstance(nested, dict) and "@id" in nested:
            out["@id"] = nested["@id"]
            if nested.get("@type", "") != "":
                out["@type"] = nested["@type"]
            del out["GEDVALUE"]

    _rename_property(out, "DATE", "dc:date")
    return out


def _simplify_event(obj: dict, old_name: str, new_class: str, label: Optional[str],
                    link_key: str, link_value: Any) -> None:
    if old_name not in obj:
        return

    def transform(el):
        if el == "Y":
            el = {}
        if not isinstance(el, dict):
            return el
        if el.get("GEDVALUE") == "Y":
            del el["GEDVALUE"]
        if link_value is not _MISSING:
            el[link_key] = link_value
        el["@type"] = new_class
        if label is not None:
            el["rdf:label"] = label
        _rename_property(el, "PLAC", "bio:place")
        return el

    _parse_property(obj, old_name, transform)
    _rename_property(obj, old_name, "bio:event")


def _simplify_individual_event(obj, old_name, new_class, label=None):
    _simplify_event(obj, old_name, new_class, label, "bio:principal", {"@id": obj["@id"]})


def _simplify_group_event(obj, old_name, new_class, label=None):
    _simplify_event(obj, old_name, new_class, label, "bio:partner",
                    obj.get("bio:participant", _MISSING))


_INDIVIDUAL_EVENTS = [
    ("BIRT", "bio:Birth", None),
    ("CHR", "bio:Baptism", None),
    ("CHRA", "bio:Baptism", None),
    ("BAPM", "bio:Baptism", None),
    ("BLES", "bio:Baptism", None),
    ("DEAT", "bio:Death", None),
    ("BURI", "bio:Burial", None),
    ("CREM", "bio:Cremation", None),
    ("ADOP", "bio:Adoption", None),
    ("BARM", "bio:BarMitzvah", None),
    ("BASM", "bio:BasMitzvah", None),
    ("CONF", "bio:IndividualEvent", "Confirmation"),
    ("FCOM", "bio:IndividualEvent", "First Communion"),
    ("ORDN", "bio:Ordination", None),
    ("NATU", "bio:Naturalization", None),
    ("EMIG", "bio:Emigration", None),
    ("IMMI", "bio:IndividualEvent", "Immigration"),
    ("PROB", "bio:IndividualEvent", "Probate"),
    ("WILL", "bio:IndividualEvent", "Will"),
    ("GRAD", "bio:Graduation", None),
    ("RETI", "bio:Retirement", None),
    ("EVEN", "bio:IndividualEvent", None),
]

_GROUP_EVENTS = [
    ("ANUL", "bio:Annulment", None),
    ("DIV", "bio:Divorce", None),
    ("DIVF", "bio:GroupEvent", "Divorce Filed"),
    ("ENGA", "bio:GroupEvent", "Engagement"),
    ("MARR", "bio:Marriage", None),
    ("MARB", "bio:GroupEvent", "Marriage Announcement"),
    ("MARC", "bio:GroupEvent", "Marriage Contract"),
    ("MARL", "bio:GroupEvent", "Marriage License"),
    ("MARS", "bio:GroupEvent", "Marriage Settlement"),
    ("EVEN", "bio:GroupEvent", None),
]


def _set_contexts(obj: dict) -> dict:
    """Renomme les propriétés selon les vocabulaires foaf / bio."""
    if obj.get("@type") == "_:INDI":
        obj["@type"] = "foaf:Person"
        _rename_property(obj, "NAME", "foaf:name")
        _rename_property(obj, "SEX", "foaf:gender")
        famc = obj.get("FAMC")
        if isinstance(famc, dict):
            fam_id = famc.pop("@id", None)
            famc["ofFamily"] = {"@id": fam_id} if fam_id is not None else {}
        for old_name, new_class, label in _INDIVIDUAL_EVENTS:
            _simplify_individual_event(obj, old_name, new_class, label)
        _rename_property(obj, "FAMS", "bio:relationship")

    elif obj.get("@type") == "_:FAM":
        obj["@type"] = "bio:Relationship"
        _rename_property(obj, "HUSB", "bio:participant")
        _rename_property(obj, "WIFE", "bio:participant")
        for old_name, new_class, label in _GROUP_EVENTS:
            _simplify_group_event(obj, old_name, new_class, label)

    return obj


def _find_spouses(fam: dict, known_id: str):
    """Trouve les conjoints d'une personne dans une famille et le type de relation."""
    participants = fam.get("bio:participant")
    if participants is None:
        return None
    if not isinstance(participants, list):
        participants = [participants]
    spouses = [p for p in participants if p.get("@id") != known_id]

    events = fam.get("bio:event")
    if events is None:
        return None
    if not isinstance(events, list):
        events = [events]
    relationship = False
    for ev in events:
        if not isinstance(ev, dict):
            continue
        if ev.get("@type") == "bio:Marriage":
            relationship = "rel:spouseOf"
            break
        if ev.get("rdf:label") == "Engagement" and relationship is False:
            relationship = "rel:engagedTo"
    # NB : comme dans la version JS, si aucun mariage/fiançailles n'est trouvé
    # la clé utilisée est "false".
    return ("false" if relationship is False else relationship), spouses


def _find_parents(fam: dict, known_id: str) -> Dict[str, list]:
    """Trouve les parents/frères et sœurs d'un enfant dans une famille.

    Comme dans la version JS, la détection pères/mères (via _FREL/_MREL) est
    désactivée : seules les fratries sont remplies.
    """
    rs: Dict[str, list] = {"fathers": [], "bioFathers": [], "mothers": [],
                           "bioMothers": [], "siblings": []}
    children = fam.get("CHIL", [])
    if not isinstance(children, list):
        children = [children]
    for child in children:
        child_id = child.get("@id") if isinstance(child, dict) else None
        if child_id != known_id:
            rs["siblings"].append({"@id": child_id})
    return rs


def _link_people(known: Dict[str, dict]) -> List[dict]:
    """Ajoute les liens conjoints / parents / fratrie entre les personnes."""
    graph_list = []
    for obj in known.values():
        if obj.get("@type") == "foaf:Person":
            # Mariages -> conjoints
            for rel in _as_array(obj, "bio:relationship"):
                fam = known.get(rel.get("@id")) if isinstance(rel, dict) else None
                if fam is None:
                    continue
                found = _find_spouses(fam, obj["@id"])
                if found is not None:
                    relationship, spouses = found
                    _set_or_add(obj, relationship, spouses)

            # Famille d'origine -> parents et fratrie
            famc = obj.get("FAMC")
            parent_fam = None
            if isinstance(famc, dict) and isinstance(famc.get("ofFamily"), dict):
                parent_fam = known.get(famc["ofFamily"].get("@id"))
            if parent_fam is not None:
                _apply_parents(known, obj, famc, _find_parents(parent_fam, obj["@id"]))

        graph_list.append(obj)
    return graph_list


def _apply_parents(known: Dict[str, dict], obj: dict, famc: dict, parents: Dict[str, list]) -> None:
    child_id = obj["@id"]

    def set_child(p):
        parent = known[p["@id"]]
        _set_or_add(parent, "bio:child", {"@id": child_id})
        if "rel:parentOf" in parent:
            _set_property(parent, "rel:parentOf",
                          [i for i in _as_array(parent, "rel:parentOf") if i.get("@id") != child_id])

    def add_birth_parent(who):
        def cb(e):
            if isinstance(e, dict) and e.get("@type") == "bio:Birth":
                _set_or_add(e, "bio:parent", who)
            return e
        return cb

    def not_in(group):
        ids = {g["@id"] for g in group}
        return lambda c: c.get("@id") not in ids

    merged = parents["fathers"] + parents["mothers"]
    if merged:
        _set_or_add(obj, "rel:childOf", merged)
        for p in merged:
            _set_or_add(known[p["@id"]], "rel:parentOf", {"@id": child_id})
    if parents["siblings"]:
        _set_or_add(obj, "rel:siblingOf", parents["siblings"])

    for group_key, prop in (("bioMothers", "bio:mother"), ("bioFathers", "bio:father")):
        group = parents[group_key]
        if group:
            _set_or_add(obj, prop, group)
            _filter_property(obj, "rel:childOf", not_in(group))
            for p in group:
                set_child(p)
            _parse_property(obj, "bio:event", add_birth_parent(group))

    pedi = famc.get("PEDI")
    if isinstance(pedi, str) and pedi.lower() in ("natural", "birth"):
        _set_or_add(obj, "bio:father", parents["fathers"])
        _set_or_add(obj, "bio:mother", parents["mothers"])
        _filter_property(obj, "rel:childOf", not_in(merged))
        _parse_property(obj, "bio:event", add_birth_parent(merged))
        for p in merged:
            set_child(p)


# ---------------------------------------------------------------------------
# API publique
# ---------------------------------------------------------------------------

def ged_to_json(ged_path) -> Dict[str, Any]:
    """Lit un fichier .ged et renvoie le graphe JSON-LD sous forme de dict."""
    graph: Dict[str, Any] = {
        "@context": {
            "foaf": "http://xmlns.com/foaf/0.1/",
            "rel": "http://purl.org/vocab/relationship",
            "bio": "http://purl.org/vocab/bio/0.1/",
            "dc": "http://purl.org/dc/elements/1.1/",
        }
    }
    known: Dict[str, dict] = {}

    for record in _ged_records(_read_lines(ged_path)):
        if isinstance(record, str):
            continue
        obj = _set_contexts(_ld_simplify(record))
        if obj.get("@id") == "_:HEAD":
            # L'en-tête devient les métadonnées du graphe
            for prop, val in obj.items():
                if prop != "@id":
                    graph[prop] = val
        elif obj.get("@id") == "_:SUBM":
            graph["SUBM"] = obj
        else:
            known[obj.get("@id")] = obj

    graph["@graph"] = _link_people(known)
    return graph


def convert_file(ged_path, json_path) -> Path:
    """Convertit ged_path en JSON et l'écrit dans json_path (UTF-8)."""
    json_path = Path(json_path)
    data = ged_to_json(ged_path)
    with open(json_path, "w", encoding="utf-8", newline="\n") as f:
        f.write(json.dumps(data, ensure_ascii=False, indent=2) + "\n")
    return json_path


if __name__ == "__main__":
    if len(sys.argv) < 2 or not sys.argv[1]:
        print("Usage : python ged_converter.py fichier.ged [sortie.json]")
        sys.exit(1)
    if len(sys.argv) >= 3:
        print(f"JSON écrit dans {convert_file(sys.argv[1], sys.argv[2])}")
    else:
        sys.stdout.reconfigure(encoding="utf-8")
        print(json.dumps(ged_to_json(sys.argv[1]), ensure_ascii=False, indent=2))
