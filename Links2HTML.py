import sys
import re
import html
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QTextEdit, QPlainTextEdit, QPushButton,
                             QFileDialog, QLabel, QMessageBox, QDialog, QCheckBox)
from PyQt6.QtGui import QIcon, QPixmap, QPainter
from PyQt6.QtCore import QByteArray, Qt, QSize, QSettings
from PyQt6.QtSvg import QSvgRenderer
from bs4 import BeautifulSoup
import mammoth
from pptx import Presentation
from pptx.enum.dml import MSO_COLOR_TYPE
from pptx.oxml.ns import qn
 
# Regular expression catching raw links (http/https) in plain text
URL_REGEX = re.compile(r'(https?://[^\s<()\"\']+)')

# Regular expression catching an *already existing* literal HTML link, e.g.
# when the user pastes raw HTML source such as <a href="...">text</a> into
# the preview area. Such links must be left untouched - they should not be
# run through URL_REGEX again (that would produce broken/nested <a> tags).
EXISTING_LINK_REGEX = re.compile(
    r'<a\s[^>]*href\s*=\s*["\'][^"\']*["\'][^>]*>.*?</a>',
    re.IGNORECASE | re.DOTALL
)

# Regular expression catching literal formatting tags (b/i/u/strong/em/sub/sup)
# that may appear as plain text (e.g. pasted raw HTML), so we know when a text
# node needs to be re-parsed as real HTML rather than left as escaped text.
FORMATTING_TAG_REGEX = re.compile(r'</?(?:b|i|u|strong|em|sub|sup)\b[^>]*>', re.IGNORECASE)

# Catches the CSS "color" property inside a style attribute. The lookbehind makes
# sure that "background-color" (or any other "*-color") is NOT matched.
COLOR_STYLE_REGEX = re.compile(r'(?<![\w-])color\s*:\s*([^;"]+)', re.IGNORECASE)


def is_color_span(tag):
    """True for the <span style="color:..."> elements created by this app."""
    return (tag.name == 'span'
            and (tag.get('style') or '').startswith('color:'))


def make_gear_icon(color="#ffffff", size=24):
    """Builds a gear icon from an inline SVG (drawn as ring + 8 teeth, so the
    centre hole is simply transparent) and returns it as a QIcon."""
    teeth = "".join(
        f'<rect x="10.5" y="1.5" width="3" height="4.5" rx="0.8" '
        f'transform="rotate({angle} 12 12)"/>'
        for angle in range(0, 360, 45)
    )
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">'
        f'<g fill="{color}">{teeth}</g>'
        f'<circle cx="12" cy="12" r="5.75" fill="none" stroke="{color}" '
        'stroke-width="3.5"/>'
        '</svg>'
    )
    renderer = QSvgRenderer(QByteArray(svg.encode("utf-8")))
    pixmap = QPixmap(size * 2, size * 2)  # 2x for crisp rendering on HiDPI
    pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    renderer.render(painter)
    painter.end()
    return QIcon(pixmap)


class SettingsDialog(QDialog):
    """Small settings window with the app options as checkboxes."""

    def __init__(self, parent, links_new_tab, detect_text_color):
        super().__init__(parent)
        self.setWindowTitle("Settings")
        self.setModal(True)
        self.setMinimumWidth(400)

        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        self.cb_new_tab = QCheckBox("Set links to be opened in new tab by default")
        self.cb_new_tab.setChecked(links_new_tab)

        self.cb_text_color = QCheckBox("Set app to recognize text color, excluding links")
        self.cb_text_color.setChecked(detect_text_color)

        layout.addWidget(self.cb_new_tab)
        layout.addWidget(self.cb_text_color)
        layout.addStretch()

        btn_close = QPushButton("Close")
        btn_close.clicked.connect(self.accept)
        layout.addWidget(btn_close)


class DocumentToHtmlConverter(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Link Converter: Word & PowerPoint -> HTML (Dark Theme)")
        self.resize(1000, 600)

        # --- Settings (both disabled by default, persisted via QSettings) ---
        # type=bool is required: QSettings may return "true"/"false" strings
        # (e.g. from an .ini backend), which would otherwise both be truthy.
        self.qsettings = QSettings("Links2HTML", "Links2HTML")
        self.links_new_tab = self.qsettings.value("links_new_tab", False, type=bool)
        self.detect_text_color = self.qsettings.value("detect_text_color", False, type=bool)
 
        self.init_ui()
        self.apply_dark_theme()
 
    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
 
        # --- Button panel (Top) ---
        btn_layout = QHBoxLayout()
 
        self.btn_load = QPushButton("Load file (Word / PowerPoint)")
        self.btn_paste = QPushButton("Paste from clipboard")
        self.btn_copy = QPushButton("Copy resulting HTML")

        # Gear button (right side) - opens the settings window
        self.btn_settings = QPushButton()
        self.btn_settings.setObjectName("settingsBtn")
        self.btn_settings.setIcon(make_gear_icon())
        self.btn_settings.setIconSize(QSize(22, 22))
        self.btn_settings.setFixedSize(40, 40)
        self.btn_settings.setToolTip("Settings")
        self.btn_settings.setCursor(Qt.CursorShape.PointingHandCursor)
 
        self.btn_load.clicked.connect(self.load_file)
        self.btn_paste.clicked.connect(self.paste_from_clipboard)
        self.btn_copy.clicked.connect(self.copy_result)
        self.btn_settings.clicked.connect(self.open_settings)
 
        btn_layout.addWidget(self.btn_load)
        btn_layout.addWidget(self.btn_paste)
        btn_layout.addWidget(self.btn_copy)
        btn_layout.addWidget(self.btn_settings)
 
        main_layout.addLayout(btn_layout)
 
        # --- Text panel (Bottom - 2 columns) ---
        text_layout = QHBoxLayout()
 
        # Left side: Preview and Drag & Drop
        left_layout = QVBoxLayout()
        left_label = QLabel("Preview (You can drop text here or load a file):")
        self.preview_area = QTextEdit()
        self.preview_area.textChanged.connect(self.process_content)
 
        left_layout.addWidget(left_label)
        left_layout.addWidget(self.preview_area)
 
        # Right side: Resulting HTML code
        right_layout = QVBoxLayout()
        right_label = QLabel("Converted HTML code:")
        self.result_area = QPlainTextEdit()
        self.result_area.setReadOnly(True) 
 
        right_layout.addWidget(right_label)
        right_layout.addWidget(self.result_area)
 
        text_layout.addLayout(left_layout)
        text_layout.addLayout(right_layout)
 
        main_layout.addLayout(text_layout)
 
    def apply_dark_theme(self):
        dark_stylesheet = """
            QMainWindow, QWidget, QDialog {
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
            QPushButton#settingsBtn {
                padding: 0px;
            }
            QLabel {
                font-size: 13px;
                font-weight: bold;
                margin-bottom: 5px;
                color: #cccccc;
            }
            QCheckBox {
                font-size: 13px;
                color: #cccccc;
                spacing: 8px;
            }
        """
        self.setStyleSheet(dark_stylesheet)

    # ------------------------------------------------------------------
    # Settings
    # ------------------------------------------------------------------
    def open_settings(self):
        dlg = SettingsDialog(self, self.links_new_tab, self.detect_text_color)
        dlg.cb_new_tab.toggled.connect(self.set_links_new_tab)
        dlg.cb_text_color.toggled.connect(self.set_detect_text_color)
        dlg.exec()

    def set_links_new_tab(self, checked):
        self.links_new_tab = checked
        self.qsettings.setValue("links_new_tab", checked)
        self.process_content()  # refresh the result immediately

    def set_detect_text_color(self, checked):
        self.detect_text_color = checked
        self.qsettings.setValue("detect_text_color", checked)
        self.process_content()  # refresh the result immediately
 
    def load_file(self):
        """Loads a .docx or .pptx file and processes it accordingly"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select file", "", "Documents (*.docx *.pptx)"
        )
        if not file_path:
            return
 
        try:
            if file_path.endswith('.docx'):
                with open(file_path, "rb") as docx_file:
                    # Map bold → <b>, italic → <i>, underline → <u>
                    # All three can coexist (mammoth wraps them independently).
                    style_map = (
                        "b => b\n"
                        "i => i\n"
                        "u => u\n"
                        "vertical-align('superscript') => sup\n"
                        "vertical-align('subscript') => sub"
                    )
                    result = mammoth.convert_to_html(docx_file, style_map=style_map)
                    self.preview_area.setHtml(result.value)
 
            elif file_path.endswith('.pptx'):
                html_content = self.parse_pptx(file_path)
                self.preview_area.setHtml(html_content)
 
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load file:\n{str(e)}")
 
    def parse_pptx(self, file_path):
        """Extracts text, hyperlinks and formatting from PowerPoint text frames"""
        prs = Presentation(file_path)
        html_output = ""

        for slide in prs.slides:
            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue

                for paragraph in shape.text_frame.paragraphs:
                    p_html = ""
                    for run in paragraph.runs:
                        text = html.escape(run.text)

                        # Wrap each formatting tag independently so they can stack:
                        # e.g. bold + underline → <b><u>text</u></b>
                        # Note: <u> is skipped for hyperlinks — browsers underline
                        # <a> elements by default, so it would be redundant.
                        font = run.font
                        is_link = bool(run.hyperlink and run.hyperlink.address)

                        # Explicit RGB text colour (links are excluded on purpose).
                        # It is always written into the preview; process_content()
                        # decides whether it ends up in the result (see Settings).
                        if not is_link:
                            try:
                                if font.color and font.color.type == MSO_COLOR_TYPE.RGB:
                                    text = f'<span style="color:#{font.color.rgb}">{text}</span>'
                            except AttributeError:
                                pass

                        # python-pptx has no direct superscript/subscript API,
                        # so read the raw <a:rPr baseline="..."/> attribute:
                        # positive baseline => superscript, negative => subscript.
                        rPr = run._r.find(qn('a:rPr'))
                        baseline = rPr.get('baseline') if rPr is not None else None
                        is_superscript = baseline is not None and int(baseline) > 0
                        is_subscript = baseline is not None and int(baseline) < 0

                        if is_superscript:
                            text = f"<sup>{text}</sup>"
                        if is_subscript:
                            text = f"<sub>{text}</sub>"
                        if font.underline and not is_link:
                            text = f"<u>{text}</u>"
                        if font.italic:
                            text = f"<i>{text}</i>"
                        if font.bold:
                            text = f"<b>{text}</b>"

                        if is_link:
                            clean_address = run.hyperlink.address.rstrip(',;.:')
                            address = html.escape(clean_address, quote=True)
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
        #QMessageBox.information(self, "Success", "HTML code copied to clipboard!")
 
    @staticmethod
    def uniform_block_color(block):
        """Returns the style string (e.g. "color:#ff0000") if ALL non-blank text
        of the block sits inside colour spans of one and the same colour.
        Returns None otherwise (mixed colours, uncoloured text, links, ...)."""
        colors = set()
        for text_node in block.find_all(string=True):
            if not text_node.strip():
                continue
            color = None
            for parent in text_node.parents:
                if parent is block:
                    break
                if is_color_span(parent):
                    color = parent['style']
                    break
            if color is None:
                return None
            colors.add(color)
        return colors.pop() if len(colors) == 1 else None

    def process_content(self):
        """Parses the preview code into clean HTML"""
        raw_html = self.preview_area.toHtml()
        soup = BeautifulSoup(raw_html, 'html.parser')

        body = soup.body if soup.body else soup

        # --- Step 1: Normalise Qt inline styles into semantic b/i/u tags ---
        # Qt encodes bold/italic/underline as inline CSS on <span> elements.
        # A single span can carry multiple styles at once, so we check all three
        # independently (no elif) and wrap the content in the appropriate tags.
        for span in body.find_all('span'):
            style = span.get('style', '')

            is_bold = ('font-weight:600' in style or 'font-weight: 600' in style or
                       'font-weight:700' in style or 'font-weight: 700' in style or
                       'font-weight:bold' in style or 'font-weight: bold' in style)
            is_italic = ('font-style:italic' in style or 'font-style: italic' in style)
            is_underline = ('text-decoration:underline' in style or
                            'text-decoration: underline' in style)
            is_superscript = ('vertical-align:super' in style or
                              'vertical-align: super' in style)
            is_subscript = ('vertical-align:sub' in style or
                            'vertical-align: sub' in style)

            # Text colour (optional feature). Spans that sit inside a link are
            # ignored, so links never get a colour of their own.
            color_value = None
            if self.detect_text_color and not span.find_parent('a'):
                color_match = COLOR_STYLE_REGEX.search(style)
                if color_match:
                    color_value = color_match.group(1).strip()

            if not (is_bold or is_italic or is_underline or
                    is_superscript or is_subscript or color_value):
                continue

            # Collect the span's children, then wrap them in layers:
            # color → sup/sub → u → i → b
            # (innermost first so the outermost tag is the first one readers see)
            children = list(span.contents)

            if color_value:
                color_tag = soup.new_tag('span')
                color_tag['style'] = f"color:{color_value}"
                for child in children:
                    color_tag.append(child)
                children = [color_tag]

            if is_superscript:
                sup_tag = soup.new_tag('sup')
                for child in children:
                    sup_tag.append(child)
                children = [sup_tag]

            if is_subscript:
                sub_tag = soup.new_tag('sub')
                for child in children:
                    sub_tag.append(child)
                children = [sub_tag]

            if is_underline:
                u_tag = soup.new_tag('u')
                for child in children:
                    u_tag.append(child.__copy__() if hasattr(child, '__copy__') else child)
                children = [u_tag]

            if is_italic:
                i_tag = soup.new_tag('i')
                for child in children:
                    i_tag.append(child)
                children = [i_tag]

            if is_bold:
                b_tag = soup.new_tag('b')
                for child in children:
                    b_tag.append(child)
                children = [b_tag]

            span.replace_with(children[0])

        # --- Step 2: Convert raw http URLs in text nodes into <a> links ---
        # Also catches literal formatting/link tags pasted as plain text
        # (e.g. raw HTML source), re-parsing them into real elements so they
        # get picked up by the strong/em/u-in-a cleanup that follows.
        for text_node in body.find_all(string=True):
            if text_node.find_parent('a'):
                continue

            original_text = str(text_node)
            if ('http' not in original_text
                    and not EXISTING_LINK_REGEX.search(original_text)
                    and not FORMATTING_TAG_REGEX.search(original_text)):
                continue

            def clean_url_match(match):
                full_url = match.group(1)
                clean_url = full_url.rstrip(',;.:')
                return f'<a href="{clean_url}">{full_url}</a>'

            # The text node may already contain literal, ready-made <a href=...>
            # markup (e.g. pasted raw HTML). Cut those chunks out first so they
            # are left completely untouched, then run URL_REGEX only on the
            # remaining plain-text pieces to catch "naked" URLs.
            existing_links = EXISTING_LINK_REGEX.findall(original_text)
            plain_pieces = EXISTING_LINK_REGEX.split(original_text)

            rebuilt = ""
            for i, piece in enumerate(plain_pieces):
                rebuilt += URL_REGEX.sub(clean_url_match, piece)
                if i < len(existing_links):
                    rebuilt += existing_links[i]

            # Re-parse unconditionally: even if URL_REGEX made no change,
            # the text node may still contain literal formatting/link tags
            # (e.g. <b>...</b>) that need to become real HTML elements.
            new_soup = BeautifulSoup(rebuilt, 'html.parser')
            text_node.replace_with(new_soup)

        # --- Step 2b: Normalise <strong>/<em> into <b>/<i> ---
        # Runs after Step 2 on purpose, so it also catches <strong>/<em> that
        # came from literal HTML text re-parsed above (not just tags that
        # mammoth/Qt produced directly).
        for strong in body.find_all('strong'):
            strong.name = 'b'
        for em in body.find_all('em'):
            em.name = 'i'

        # --- Step 2c: Strip redundant <u> inside <a> ---
        # Browsers underline anchor elements by default, so <u> inside <a> is
        # purely noise. Runs after Step 2 so it also catches <u> that came
        # from literal HTML text re-parsed above. Unwrap every <u> that has
        # an <a> ancestor.
        for u_tag in body.find_all('u'):
            if u_tag.find_parent('a'):
                u_tag.unwrap()

        # --- Step 2d: Keep colour away from links ---
        # Step 2 may have created <a> tags (naked URLs) inside a coloured span.
        # In that case the colour is re-applied to each text piece outside the
        # link separately, and the original wrapping span is removed.
        if self.detect_text_color:
            for color_span in body.find_all(is_color_span):
                if color_span.find('a') is None:
                    continue
                color_style = color_span['style']
                for text_node in color_span.find_all(string=True):
                    if text_node.find_parent('a') or not text_node.strip():
                        continue
                    wrapper = soup.new_tag('span')
                    wrapper['style'] = color_style
                    text_node.replace_with(wrapper)
                    wrapper.append(text_node)
                color_span.unwrap()

        # --- Step 2e: Open links in a new tab (optional feature) ---
        if self.links_new_tab:
            for a_tag in body.find_all('a'):
                a_tag['target'] = '_blank'

        # --- Step 3: Build clean output — keep <a>, <b>, <i>, <u> ---
        # (plus colour spans created above, when the colour option is enabled)
        KEEP_TAGS = {'a', 'b', 'i', 'u', 'sub', 'sup'}
        result_text = ""
        for block in body.find_all(['p', 'div', 'li', 'h1', 'h2', 'h3']):
            # If the WHOLE paragraph has one colour, the colour is moved from
            # the <span>s to the paragraph itself: <p style="color:#...">...</p>.
            # Mixed/partial colouring keeps inline <span>s (a <p> can't be
            # placed in the middle of a sentence).
            paragraph_style = None
            if self.detect_text_color and block.name in ('p', 'div'):
                paragraph_style = self.uniform_block_color(block)
                if paragraph_style:
                    for color_span in block.find_all(is_color_span):
                        color_span.unwrap()

            tags_to_unwrap = [tag for tag in block.find_all(True)
                              if tag.name not in KEEP_TAGS and not is_color_span(tag)]
            for tag in tags_to_unwrap:
                tag.unwrap()

            block_html = "".join(str(c) for c in block.contents).strip()
            if block_html:
                if paragraph_style:
                    block_html = f'<p style="{html.escape(paragraph_style, quote=True)}">{block_html}</p>'
                result_text += block_html + "\n\n"

        self.result_area.setPlainText(result_text.strip())
 
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DocumentToHtmlConverter()
    window.show()
    sys.exit(app.exec())
