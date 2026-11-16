import sys
from pathlib import Path
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel,
    QFileDialog, QLineEdit, QHBoxLayout, QRadioButton, QStackedWidget, QCheckBox, 
)

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QHBoxLayout
from generator import generate_docx

PIXEL_PER_CHAR = 10
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
        h_box.addLayout(layout1)
        h_box.addLayout(layout2)
        # self.setGeometry(300,300)
        # layout1.resize(window.width()//2, window.height()//2)
        # tool_bar = QToolBar()
        # tool_bar.addWidget(tool_bar)
        # layout.addLayout(tool_bar)

        # JSON input
        h_json = QHBoxLayout()
        # h_json.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.json_input = QLineEdit(str(Path(__file__).resolve().parent / 'Hrysio3.json'))
        btn_json = QPushButton('Browse JSON')
        btn_json.clicked.connect(self.browse_json)
        h_json.width = 200
        h_json.addWidget(self.json_input)
        h_json.addWidget(btn_json)
        layout1.addLayout(h_json)

        # Bouton vers la page Index
        btn_indexs = QPushButton("Index")
        btn_indexs.clicked.connect(self.window_.show_index)
        layout1.addWidget(btn_indexs)

        # Output
        h_out = QHBoxLayout()
        text_area = QLineEdit('output')
        self.out_input = text_area
        self.out_input.setFixedWidth(100)
        btn_out = QPushButton('Browse Output')
        btn_out.clicked.connect(self.browse_out)
        btn_out.setFixedWidth(len(btn_out.text())*PIXEL_PER_CHAR)
        h_out.setAlignment(Qt.AlignRight)
        h_out.addWidget(self.out_input)
        h_out.addWidget(btn_out)
        layout1.addLayout(h_out)


        # Generate button
        self.btn_generate = QPushButton('Generate')
        self.btn_generate.clicked.connect(self.generate)
        layout1.addWidget(self.btn_generate)
        self.btn_test1 = QRadioButton("Activer index ?")
        # self.btn_test1.clicked.connect()
        # self.btn_test2 = QRadioButton("test2")
        self.btn_test1.clicked.connect(self.test)
        layout2.addWidget(self.btn_test1)
        # layout2.addLayout(self.btn_test1)
        # layout.addWidget(self.btn_test2)

        # Status
        self.status = QLabel('')
        layout1.addWidget(self.status)

        self.setLayout(h_box)
        # self.addLayout(layout2)


    def browse_json(self):
        path, _ = QFileDialog.getOpenFileName(self, 'Open JSON', '', 'JSON Files (*.json);;All Files (*)')
        if path:
            self.json_input.setText(path)

    def browse_out(self,text_area):
        path, _ = QFileDialog.getSaveFileName(self, 'Save DOCX',text_area.text(), 'Word Documents (*.docx)')
        if path:
            self.out_input.setText(path)

    def generate(self):
        self.status.setText('Generating...')
        QApplication.processEvents()
        try:
            generate_docx(json_path=self.json_input.text(), output_path=self.out_input.text()+'.docx',
                          indexNoms=self.window_.index_page.btn_noms.isChecked(),
                          indexLieux=self.window_.index_page.btn_lieux.isChecked(),
                          indexProfessions=self.window_.index_page.btn_professions.isChecked())
            self.status.setText(f'Done: {self.out_input.text()}')
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




if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())
