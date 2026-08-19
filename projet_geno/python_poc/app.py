import sys
from pathlib import Path
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel,
    QFileDialog, QLineEdit, QHBoxLayout, QRadioButton
)

from generator import generate_docx


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Document Generator (PySide6)')
        self.resize(600, 150)
        layout = QVBoxLayout()

        # JSON input
        h_json = QHBoxLayout()
        self.json_input = QLineEdit(str(Path(__file__).resolve().parent / 'test.json'))
        btn_json = QPushButton('Browse JSON')
        btn_json.clicked.connect(self.browse_json)
        h_json.addWidget(self.json_input)
        h_json.addWidget(btn_json)
        layout.addLayout(h_json)

        # Output
        h_out = QHBoxLayout()
        self.out_input = QLineEdit(str(Path(__file__).resolve().parent / 'output.docx'))
        btn_out = QPushButton('Browse Output')
        btn_out.clicked.connect(self.browse_out)
        h_out.addWidget(self.out_input)
        h_out.addWidget(btn_out)
        layout.addLayout(h_out)


        # Generate button
        self.btn_generate = QPushButton('Generate')
        self.btn_generate.clicked.connect(self.generate)
        layout.addWidget(self.btn_generate)
        self.btn_test1 = QRadioButton("test1")
        # self.btn_test2 = QRadioButton("test2")
        self.btn_test1.clicked.connect(self.test)
        layout.addWidget(self.btn_test1)
        # layout.addWidget(self.btn_test2)

        # Status
        self.status = QLabel('')
        layout.addWidget(self.status)

        self.setLayout(layout)

    def browse_json(self):
        path, _ = QFileDialog.getOpenFileName(self, 'Open JSON', '', 'JSON Files (*.json);;All Files (*)')
        if path:
            self.json_input.setText(path)

    def browse_out(self):
        path, _ = QFileDialog.getSaveFileName(self, 'Save DOCX', str(Path.cwd() / 'output.docx'), 'Word Documents (*.docx)')
        if path:
            self.out_input.setText(path)

    def generate(self):
        self.status.setText('Generating...')
        QApplication.processEvents()
        try:
            generate_docx(self.json_input.text(), self.out_input.text())
            self.status.setText(f'Done: {self.out_input.text()}')
        except Exception as e:
            self.status.setText(f'Error: {e}')

    def test(self):
        if(self.btn_test1.isChecked() == True):
            print("test")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())
