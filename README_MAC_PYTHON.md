# GED2DOCX pour Mac — à partir du code source (avec Python)

Ce guide est pour vous si vous êtes à l'aise avec le Terminal et voulez lancer l'application depuis le code (pour la modifier, tester une version en cours, etc.).

> Vous voulez juste utiliser l'application, sans rien installer ? Suivez plutôt **[README_MAC.md](README_MAC.md)**.

## Prérequis

- **Python 3.10 ou plus récent.** Vérifiez avec :
  ```bash
  python3 --version
  ```
  S'il manque, installez-le depuis [python.org/downloads/macos](https://www.python.org/downloads/macos/) ou avec Homebrew (`brew install python`).
- **git** (facultatif) : `xcode-select --install` l'installe. Sans git, vous pouvez télécharger le code en ZIP (bouton **Code → Download ZIP** sur GitHub).

Node.js n'est **pas** nécessaire : la conversion GED → JSON est faite en Python (`projet_geno/python_poc/ged_converter.py`).

## 1. Récupérer le code

```bash
git clone https://github.com/mohawk2409/genopresse-project.git
cd genopresse-project
```

## 2. Lancer l'application

### Option A — double-clic

Double-cliquez sur **`lancer_mac.command`** à la racine du projet.
Au premier lancement, il crée un environnement virtuel `.venv` et installe les dépendances (1 à 2 minutes), puis ouvre l'application.

Si macOS refuse de l'ouvrir (« permission refusée » ou fichier bloqué), une seule fois dans le Terminal :
```bash
chmod +x lancer_mac.command
xattr -d com.apple.quarantine lancer_mac.command 2>/dev/null
```

### Option B — Terminal

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python projet_geno/python_poc/app.py
```

Les fichiers générés (JSON, rapport intermédiaire, `.docx`) sont enregistrés dans **`projet_geno/target/`**.

### Vérification rapide sans interface

```bash
QT_QPA_PLATFORM=offscreen python projet_geno/python_poc/app.py --selftest projet_geno/Hrysio-Export.ged
```
Doit se terminer par `selftest: conversion + génération OK`.

### Convertir un .ged en ligne de commande

```bash
python projet_geno/python_poc/ged_converter.py mon-fichier.ged sortie.json
```

## 3. Construire l'application Mac (.app) soi-même

Sur un Mac, depuis l'environnement virtuel :
```bash
bash packaging/build_mac.sh
```
Résultat : `dist/GED2DOCX.app` et `dist/GED2DOCX-mac-<apple-silicon|intel>.zip`.
L'application est construite pour l'architecture du Mac utilisé (Apple Silicon ou Intel).

## 4. La pipeline GitHub (pour les mainteneurs)

Le fichier [`.github/workflows/build-mac.yml`](.github/workflows/build-mac.yml) construit automatiquement l'application sur des Mac de GitHub :

| Événement | Ce qui se passe |
|---|---|
| Push sur `main` | Construit et teste les 2 versions (Apple Silicon + Intel). Les zips sont téléchargeables dans l'onglet **Actions** → le run → **Artifacts**. |
| Push d'un tag `v*` | Pareil, **et** publie une **Release** GitHub avec les 2 zips — c'est le lien donné aux utilisateurs dans `README_MAC.md`. |
| Onglet Actions → **Run workflow** | Lance une construction à la main. |

Chaque construction exécute `--selftest` sur l'application finale : si l'app ne démarre pas ou ne génère pas de document, la pipeline échoue et rien n'est publié.

**Publier une nouvelle version :**
```bash
git tag v1.0.0
git push origin v1.0.0
```

**Points d'attention :**
- Le dépôt doit être **public** pour que les utilisateurs puissent télécharger la Release sans compte GitHub.
- L'application n'est **pas signée** par un compte développeur Apple (99 $/an), d'où l'étape « Ouvrir quand même » au premier lancement. Pour la supprimer, il faudrait signer et notariser l'app (`codesign` + `xcrun notarytool`) dans la pipeline.
- Le runner Intel (`macos-15-intel`) est le dernier proposé par GitHub pour Intel. S'il disparaît, supprimez simplement sa ligne dans le workflow.
