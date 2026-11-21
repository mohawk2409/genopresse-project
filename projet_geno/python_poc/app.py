import sys
from pathlib import Path
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel,
    QFileDialog, QLineEdit, QHBoxLayout, QRadioButton, QStackedWidget, QCheckBox,
    QGridLayout,
)

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from PySide6.QtWidgets import QHBoxLayout
from generator import generate_docx
from ged_converter import convert_file



PIXEL_PER_CHAR = 10
BUTTON_WIDTH = 160                                     # même largeur pour tous les boutons "Browse"
PROJECT_DIR = Path(__file__).resolve().parent.parent   # dossier projet_geno

# Application empaquetée (.app Mac / .exe créé par PyInstaller) : on ne peut pas écrire
# à l'intérieur de l'application, donc les fichiers générés vont dans Documents/GED2DOCX.
IS_PACKAGED = getattr(sys, 'frozen', False)
if IS_PACKAGED:
    TARGET_DIR = Path.home() / 'Documents' / 'GED2DOCX'
    DEFAULT_GED = ''
else:
    TARGET_DIR = PROJECT_DIR / 'target'                 # dossier où sont enregistrés les JSON convertis
    DEFAULT_GED = str(PROJECT_DIR / 'Hrysio-Export.ged')

class MainWindow(QWidget):
    """Fenêtre unique qui contient toutes les pages et permet de passer de l'une à l'autre."""
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Document Generator (PySide6)')
        self.resize(600, 300)

        self.stack = QStackedWidget()
        self.main_page = MainPage(self)
        self.index_page = IndexPage(self)
        self.stack.addWidget(self.main_page)
        self.stack.addWidget(self.index_page)
        h_box = QHBoxLayout()
        layout1 = QVBoxLayout()
        layout1.addWidget(self.stack)
        self.setLayout(layout1)
        layout2 = QVBoxLayout()
        self.setLayout(layout2)
        # h_box.addLayout(layout1)
        # h_box.addLayout(layout2)
    def show_main(self):
        self.stack.setCurrentWidget(self.main_page)

    def show_index(self):
        self.stack.setCurrentWidget(self.index_page)


class MainPage(QWidget):
    def __init__(self, window):
        super().__init__()
        self.window_ = window
        h_box = QHBoxLayout()
        layout1 = QVBoxLayout()
        layout2 = QVBoxLayout()
        h_box.addLayout(layout1, 1)   # la colonne principale prend toute la largeur disponible
        h_box.addLayout(layout2)
        # self.setGeometry(300,300)
        # layout1.resize(window.width()//2, window.height()//2)
        # tool_bar = QToolBar()
        # tool_bar.addWidget(tool_bar)
        # layout.addLayout(tool_bar)

        # Grille : colonne 0 = libellés, colonne 1 = champs (s'étirent), colonne 2 = boutons
        grid = QGridLayout()
        grid.setColumnStretch(1, 1)
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(6)
        label_align = Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter

        def add_row(row, label, field, button):
            grid.addWidget(QLabel(label), row, 0, alignment=label_align)
            grid.addWidget(field, row, 1)
            button.setFixedWidth(BUTTON_WIDTH)
            grid.addWidget(button, row, 2)

        # Ligne 0 : fichier GED source
        self.ged_input = QLineEdit(DEFAULT_GED)
        self.ged_input.setPlaceholderText('Choisis un fichier .ged avec "Browse GED"')
        btn_ged = QPushButton('Browse GED')
        btn_ged.clicked.connect(self.browse_ged)
        title = QLabel("GED2DOCX")
        title.setFont(QFont('Arial', 24, QFont.Weight.Bold))   # taille 24, en gras
        title.setContentsMargins(0, 0, 0, 12)                   # un peu d'espace sous le titre
        # ligne 0, colonne 0, sur 1 ligne et 3 colonnes -> centré sur toute la largeur
        grid.addWidget(title, 0, 0, 1, 3, alignment=Qt.AlignmentFlag.AlignCenter)
        add_row(1, 'Fichier GED :', self.ged_input, btn_ged)

        # Ligne 1 : nom du JSON à générer + conversion
        self.json_name_input = QLineEdit('report')
        self.json_name_input.setPlaceholderText('Nom du fichier JSON')
        btn_convert = QPushButton('Convert GED → JSON')
        btn_convert.clicked.connect(self.convert_ged)
        add_row(2, 'Nom du JSON :', self.json_name_input, btn_convert)

        # Ligne 2 : JSON utilisé pour la génération (= target/<nom>.json, mis à jour en direct)
        self.json_input = QLineEdit()
        btn_json = QPushButton('Browse JSON')
        btn_json.clicked.connect(self.browse_json)
        add_row(3, 'Fichier JSON :', self.json_input, btn_json)
        self.json_name_input.textChanged.connect(self.update_json_path)
        self.update_json_path(self.json_name_input.text())

        # Ligne 3 : fichier Word de sortie
        self.out_input = QLineEdit('output')
        btn_out = QPushButton('Browse Output')
        btn_out.clicked.connect(self.browse_out)
        add_row(4, 'Sortie DOCX :', self.out_input, btn_out)

        layout1.addLayout(grid)
        layout1.addSpacing(10)

        # Boutons d'action côte à côte
        h_actions = QHBoxLayout()
        btn_indexs = QPushButton("Index")
        btn_indexs.clicked.connect(self.window_.show_index)
        self.btn_generate = QPushButton('Generate')
        self.btn_generate.clicked.connect(self.generate)
        self.btn_generate.setDefault(True)
        h_actions.addWidget(btn_indexs)
        h_actions.addStretch()
        h_actions.addWidget(self.btn_generate)
        layout1.addLayout(h_actions)
        self.btn_test1 = QRadioButton("Activer index ?")
        # self.btn_test1.clicked.connect()
        # self.btn_test2 = QRadioButton("test2")
        self.btn_test1.clicked.connect(self.test)
        layout2.addWidget(self.btn_test1)
        # layout2.addLayout(self.btn_test1)
        # layout.addWidget(self.btn_test2)

        # Status
        self.status = QLabel('')
        self.status.setWordWrap(True)
        layout1.addWidget(self.status)
        layout1.addStretch()

        self.setLayout(h_box)
        # self.addLayout(layout2)


    def browse_ged(self):
        start_dir = str(Path.home()) if IS_PACKAGED else str(PROJECT_DIR)
        path, _ = QFileDialog.getOpenFileName(self, 'Open GED', start_dir, 'GEDCOM Files (*.ged);;All Files (*)')
        if path:
            self.ged_input.setText(path)

    def target_json_path(self):
        """Chemin target/<nom>.json à partir du nom saisi (None si le nom est vide)."""
        name = self.json_name_input.text().strip()
        if not name:
            return None
        if not name.lower().endswith('.json'):
            name += '.json'
        return TARGET_DIR / name

    def update_json_path(self, _text=None):
        """Met à jour le champ 'Fichier JSON' à chaque modification du nom."""
        path = self.target_json_path()
        self.json_input.setText(str(path) if path else '')

    def convert_ged(self):
        ged_path = Path(self.ged_input.text().strip())
        json_path = self.target_json_path()
        if not ged_path.is_file():
            self.status.setText(f'Error: fichier GED introuvable : {ged_path}')
            return
        if json_path is None:
            self.status.setText('Error: indique un nom pour le fichier JSON')
            return

        self.status.setText('Converting...')
        QApplication.processEvents()
        try:
            TARGET_DIR.mkdir(parents=True, exist_ok=True)
            convert_file(ged_path, json_path)
            self.json_input.setText(str(json_path))   # le JSON est prêt pour "Generate"
            self.status.setText(f'Converted: {json_path}')
        except Exception as e:
            self.status.setText(f'Error: {e}')

    def browse_json(self):
        path, _ = QFileDialog.getOpenFileName(self, 'Open JSON', '', 'JSON Files (*.json);;All Files (*)')
        if path:
            self.json_input.setText(path)

    def browse_out(self):
        path, _ = QFileDialog.getSaveFileName(self, 'Save DOCX', self.out_input.text(), 'Word Documents (*.docx)')
        if path:
            self.out_input.setText(path)

    def generate(self):
        self.status.setText('Generating...')
        QApplication.processEvents()
        out = Path(self.out_input.text().strip() or 'output')
        if out.suffix.lower() != '.docx':
            out = out.with_name(out.name + '.docx')
        if not out.is_absolute():
            out = TARGET_DIR / out    # un simple nom ("output") -> enregistré dans le dossier cible
        try:
            TARGET_DIR.mkdir(parents=True, exist_ok=True)
            out.parent.mkdir(parents=True, exist_ok=True)
            generate_docx(json_path=self.json_input.text(), output_path=str(out),
                          report_path=str(TARGET_DIR / 'extracted_report.json'),
                          indexNoms=self.window_.index_page.btn_noms.isChecked(),
                          indexLieux=self.window_.index_page.btn_lieux.isChecked(),
                          indexProfessions=self.window_.index_page.btn_professions.isChecked())
            self.status.setText(f'Done: {out}')
        except Exception as e:
            self.status.setText(f'Error: {e}')

    def test(self):
        if(self.btn_test1.isChecked() == True):
            print(self.btn_test1.isChecked())
    


class IndexPage(QWidget):
    def __init__(self, window):
        super().__init__()
        self.window_ = window
        layout = QVBoxLayout()
        btn_back = QPushButton("← Retour")
        btn_back.clicked.connect(self.window_.show_main)
        layout.addWidget(btn_back)

        # Ajoute tes boutons ici
        self.btn_noms = QCheckBox("Index des noms")
        self.btn_lieux = QCheckBox("Index des lieux")
        self.btn_professions = QCheckBox("Index des professions")
        layout.addWidget(self.btn_noms, alignment=Qt.AlignHCenter)
        layout.addWidget(self.btn_lieux, alignment=Qt.AlignHCenter)
        layout.addWidget(self.btn_professions, alignment=Qt.AlignHCenter)

        layout.addStretch()
        self.setLayout(layout)




def selftest(ged_path=None) -> int:
    """Vérification automatique (utilisée par la pipeline GitHub) : `app --selftest [fichier.ged]`.

    Crée la fenêtre sans l'afficher, puis si un .ged est fourni, fait la conversion
    et la génération complètes dans un dossier temporaire.
    """
    import tempfile
    app = QApplication.instance() or QApplication(sys.argv)
    MainWindow()
    print('selftest: fenêtre OK')
    if ged_path:
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            convert_file(ged_path, tmp / 'test.json')
            out = generate_docx(json_path=str(tmp / 'test.json'),
                                report_path=str(tmp / 'report.json'),
                                output_path=str(tmp / 'test.docx'))
            size = Path(out).stat().st_size
            print(f'selftest: conversion + génération OK ({size} octets)')
    app.quit()
    return 0


if __name__ == '__main__':
    if '--selftest' in sys.argv:
        args = [a for a in sys.argv[1:] if a != '--selftest']
        sys.exit(selftest(args[0] if args else None))

    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())
