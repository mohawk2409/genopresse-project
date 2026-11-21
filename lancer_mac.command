#!/bin/bash
# Lance GED2DOCX sur Mac à partir du code source (Python 3 requis).
# Double-cliquez sur ce fichier dans le Finder.
# Au premier lancement, il crée un environnement Python (.venv) et installe les dépendances.
set -e
cd "$(dirname "$0")"

if ! command -v python3 >/dev/null 2>&1; then
    echo "Python 3 n'est pas installé. Téléchargez-le sur https://www.python.org/downloads/macos/"
    read -r -p "Appuyez sur Entrée pour fermer…"
    exit 1
fi

if [ ! -d .venv ]; then
    echo "Premier lancement : installation (1 à 2 minutes)…"
    python3 -m venv .venv
fi
source .venv/bin/activate
python -m pip install --quiet --upgrade pip
python -m pip install --quiet -r requirements.txt

python projet_geno/python_poc/app.py
