from docx import Document
from docx.shared import Inches


def test():
    document = Document() # Création d'un nouveau document

    document.add_heading('Document Title', 0)


    document.save('demo.docx')


if __name__ == "__main__":
    test()