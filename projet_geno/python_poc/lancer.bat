@echo off

REM === 1. Lancer le script Python ===
echo Lancement du script Python...
python generator.py --json Hrysio2.json --report report.json --out new_output.docx

REM === 2. Ouvrir le fichier Word ===
echo Ouverture du fichier Word...
start "" "new_output.docx"

echo Termine !
pause
