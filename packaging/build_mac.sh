#!/bin/bash
# Construit l'application Mac GED2DOCX.app avec PyInstaller, puis la zippe.
#
# Utilisé par la pipeline GitHub (.github/workflows/build-mac.yml),
# mais peut aussi être lancé à la main sur un Mac :
#     bash packaging/build_mac.sh
#
# Résultat : dist/GED2DOCX.app  et  dist/GED2DOCX-mac-<arch>.zip
set -euo pipefail

cd "$(dirname "$0")/.."          # racine du dépôt
APP_NAME="GED2DOCX"
SRC_DIR="projet_geno/python_poc"
ARCH="$(uname -m)"               # arm64 (Apple Silicon) ou x86_64 (Intel)

echo "==> Installation des dépendances"
python3 -m pip install --upgrade pip
python3 -m pip install -r requirements.txt pyinstaller

echo "==> Construction de ${APP_NAME}.app (${ARCH})"
rm -rf build dist
DOCX_DIR="$(python3 -c 'import docx, os; print(os.path.dirname(docx.__file__))')"
python3 -m PyInstaller \
    --noconfirm \
    --windowed \
    --name "${APP_NAME}" \
    --osx-bundle-identifier "com.genopresse.ged2docx" \
    --paths "${SRC_DIR}" \
    --collect-data docx \
	--add-data "${DOCX_DIR}/parts:docx/parts" \
    "${SRC_DIR}/app.py"

echo "==> Vérification : l'application démarre et génère un document"
QT_QPA_PLATFORM=offscreen "dist/${APP_NAME}.app/Contents/MacOS/${APP_NAME}" \
    --selftest "projet_geno/Hrysio-Export.ged"

echo "==> Création du zip"
case "${ARCH}" in
    arm64)  LABEL="apple-silicon" ;;
    x86_64) LABEL="intel" ;;
    *)      LABEL="${ARCH}" ;;
esac
ZIP="dist/${APP_NAME}-mac-${LABEL}.zip"
# ditto conserve la structure et les attributs du .app (contrairement à "zip")
ditto -c -k --sequesterRsrc --keepParent "dist/${APP_NAME}.app" "${ZIP}"

echo "==> Terminé : ${ZIP}"
