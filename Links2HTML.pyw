import sys
import re
import html
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QTextEdit, QPlainTextEdit, QPushButton,
                             QFileDialog, QLabel, QMessageBox)
from bs4 import BeautifulSoup
import mammoth
from pptx import Presentation

# Wyrażenie regularne wyłapujące surowe linki (http/https) w zwykłym tekście
URL_REGEX = re.compile(r'(https?://[^\s<()\"\']+)')

class DocumentToHtmlConverter(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Konwerter Linków: Word & PowerPoint -> HTML (Dark Theme)")
        self.resize(1000, 600)
        
        self.init_ui()
        self.apply_dark_theme()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # --- Panel przycisków (Góra) ---
        btn_layout = QHBoxLayout()
        
        self.btn_load = QPushButton("Wczytaj plik (Word / PowerPoint)")
        self.btn_paste = QPushButton("Wklej ze schowka")
        self.btn_copy = QPushButton("Kopiuj wynikowy HTML")

        self.btn_load.clicked.connect(self.load_file)
        self.btn_paste.clicked.connect(self.paste_from_clipboard)
        self.btn_copy.clicked.connect(self.copy_result)

        btn_layout.addWidget(self.btn_load)
        btn_layout.addWidget(self.btn_paste)
        btn_layout.addWidget(self.btn_copy)
        
        main_layout.addLayout(btn_layout)

        # --- Panel tekstowy (Dół - 2 kolumny) ---
        text_layout = QHBoxLayout()

        # Lewa strona: Podgląd i Drag & Drop
        left_layout = QVBoxLayout()
        left_label = QLabel("Podgląd (Możesz tu upuścić tekst lub wczytać plik):")
        self.preview_area = QTextEdit()
        self.preview_area.textChanged.connect(self.process_content)
        
        left_layout.addWidget(left_label)
        left_layout.addWidget(self.preview_area)

        # Prawa strona: Wynikowy kod HTML
        right_layout = QVBoxLayout()
        right_label = QLabel("Skonwertowany kod HTML:")
        self.result_area = QPlainTextEdit()
        self.result_area.setReadOnly(True) 
        
        right_layout.addWidget(right_label)
        right_layout.addWidget(self.result_area)

        text_layout.addLayout(left_layout)
        text_layout.addLayout(right_layout)
        
        main_layout.addLayout(text_layout)

    def apply_dark_theme(self):
        dark_stylesheet = """
            QMainWindow, QWidget {
                background-color: #2b2b2b;
                color: #a9b7c6;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
            QTextEdit, QPlainTextEdit {
                background-color: #1e1e1e;
                color: #a9b7c6;
                border: 1px solid #555555;
                border-radius: 6px;
                padding: 8px;
                font-size: 14px;
            }
            QPushButton {
                background-color: #365880;
                color: #ffffff;
                border: none;
                border-radius: 6px;
                padding: 10px 15px;
                font-weight: bold;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #436a9b;
            }
            QPushButton:pressed {
                background-color: #27405e;
            }
            QLabel {
                font-size: 13px;
                font-weight: bold;
                margin-bottom: 5px;
                color: #cccccc;
            }
        """
        self.setStyleSheet(dark_stylesheet)

    def load_file(self):
        """Wczytuje plik .docx lub .pptx i odpowiednio go procesuje"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Wybierz plik", "", "Dokumenty (*.docx *.pptx)"
        )
        if not file_path:
            return

        try:
            if file_path.endswith('.docx'):
                with open(file_path, "rb") as docx_file:
                    result = mammoth.convert_to_html(docx_file)
                    self.preview_area.setHtml(result.value)
            
            elif file_path.endswith('.pptx'):
                html_content = self.parse_pptx(file_path)
                self.preview_area.setHtml(html_content)
                
        except Exception as e:
            QMessageBox.critical(self, "Błąd", f"Nie udało się wczytać pliku:\n{str(e)}")

    def parse_pptx(self, file_path):
        """Wyciąga tekst i hiperłącza z pól tekstowych prezentacji PowerPoint"""
        prs = Presentation(file_path)
        html_output = ""

        # Przelatujemy przez wszystkie slajdy
        for slide in prs.slides:
            # Przelatujemy przez wszystkie kształty (pola tekstowe, tabele itp.)
            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue
                
                # Zbieramy paragrafy z pola tekstowego
                for paragraph in shape.text_frame.paragraphs:
                    p_html = ""
                    for run in paragraph.runs:
                        text = html.escape(run.text)
                        # Sprawdzamy czy dany fragment tekstu ma podpięte hiperłącze
                        if run.hyperlink and run.hyperlink.address:
                            # ZABEZPIECZENIE: Ucieczka znaków specjalnych w URL (np. cudzysłowów)
                            address = html.escape(run.hyperlink.address, quote=True)
                            p_html += f'<a href="{address}">{text}</a>'
                        else:
                            p_html += text
                    
                    if p_html.strip():
                        html_output += f"<p>{p_html}</p>\n"
                        
        return html_output

    def paste_from_clipboard(self):
        self.preview_area.paste()

    def copy_result(self):
        QApplication.clipboard().setText(self.result_area.toPlainText())
        QMessageBox.information(self, "Sukces", "Kod HTML został skopiowany do schowka!")

    def process_content(self):
        """Parsuje kod z podglądu na czysty HTML"""
        raw_html = self.preview_area.toHtml()
        soup = BeautifulSoup(raw_html, 'html.parser')
        
        body = soup.body if soup.body else soup

        # Zamiana surowych linków tekstowych na tagi <a>
        for text_node in body.find_all(string=True):
            # POPRAWKA 1: Sprawdzamy wszystkich przodków, bo QTextEdit dodaje tagi <span>
            if text_node.find_parent('a'):
                continue

            original_text = str(text_node)
            if 'http' in original_text:
                replaced_text = URL_REGEX.sub(r'<a href="\1">\1</a>', original_text)
                if replaced_text != original_text:
                    new_soup = BeautifulSoup(replaced_text, 'html.parser')
                    text_node.replace_with(new_soup)

        # Oczyszczanie struktury - zostawiamy tylko same linki wewnątrz paragrafów
        result_text = ""
        for block in body.find_all(['p', 'div', 'li', 'h1', 'h2', 'h3']):
            # POPRAWKA 2: Zbieramy tagi do listy, aby bezpiecznie modyfikować drzewo
            tags_to_unwrap = [tag for tag in block.find_all(True) if tag.name != 'a']
            for tag in tags_to_unwrap:
                tag.unwrap()

            block_html = "".join(str(c) for c in block.contents).strip()
            if block_html:
                result_text += block_html + "\n\n"

        self.result_area.setPlainText(result_text.strip())

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DocumentToHtmlConverter()
    window.show()
    sys.exit(app.exec())