import os
import sys
import re
import html
import tempfile
from functools import partial
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QTextEdit, QPlainTextEdit, QPushButton,
                             QFileDialog, QLabel, QMessageBox, QDialog, QCheckBox)
from PyQt6.QtGui import QIcon, QPixmap, QPainter, QTextCursor, QColor, QImage
from PyQt6.QtCore import QByteArray, Qt, QSize, QSettings
from PyQt6.QtSvg import QSvgRenderer
from bs4 import BeautifulSoup
import mammoth
from pptx import Presentation
from pptx.enum.dml import MSO_COLOR_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
 
# Folder of the script (or of the .exe when frozen) - the "icons" folder lives next to it
APP_DIR = os.path.dirname(sys.executable if getattr(sys, "frozen", False)
                          else os.path.abspath(__file__))
APP_ICON_PATH = os.path.join(APP_DIR, "icons", "Link2HTML(dark).ico")

# Text colour used in the first (preview) window, regardless of the pasted content
PREVIEW_TEXT_COLOR = "#a0a0a0"

# What an empty line is converted to (when the option is enabled)
EMPTY_LINE_HTML = "<br><br>"

# All options shown as checkboxes in the Settings window:
# (QSettings key, checkbox label, default value)
SETTINGS_OPTIONS = [
    ("links_new_tab", "Set links to be opened in new tab by default", False),
    ("detect_text_color", "Set app to recognize text color, excluding links", False),
    ("detect_background", "Set app to recognize text background (highlight), excluding links", False),
    ("detect_strikethrough", "Preserve strikethrough text (<s>)", False),
    ("keep_headings", "Preserve headings (<h1> - <h6>)", False),
    ("keep_lists", "Preserve bulleted / numbered lists (<ul>, <ol>, <li>)", False),
    ("keep_alignment", "Preserve paragraph alignment (center / right / justify)", False),
    ("empty_line_br", f"Convert empty lines to {EMPTY_LINE_HTML}", True),
]

HEADING_TAGS = ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']


def load_app_icon():
    """Returns the application icon from ./icons/Link2HTML(dark).ico
    (an empty QIcon when the file is missing)."""
    if not os.path.isfile(APP_ICON_PATH):
        print(f"Warning: icon not found: {APP_ICON_PATH}", file=sys.stderr)
        return QIcon()
    return QIcon(APP_ICON_PATH)


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
FORMATTING_TAG_REGEX = re.compile(r'</?(?:b|i|u|s|strike|del|strong|em|sub|sup)\b[^>]*>', re.IGNORECASE)

# Literal formatting tags (b/i/u/s/strike/del/strong/em/sub/sup) - see above.
# Catches the CSS "color" property inside a style attribute. The lookbehind makes
# sure that "background-color" (or any other "*-color") is NOT matched.
COLOR_STYLE_REGEX = re.compile(r'(?<![\w-])color\s*:\s*([^;"]+)', re.IGNORECASE)


# Catches the CSS "background-color" property inside a style attribute.
BACKGROUND_STYLE_REGEX = re.compile(r'background-color\s*:\s*([^;"]+)', re.IGNORECASE)

# Catches the "text-decoration" value (may hold "underline line-through" etc.)
TEXT_DECORATION_REGEX = re.compile(r'text-decoration\s*:\s*([^;"]+)', re.IGNORECASE)

# Catches paragraph alignment (the lookbehind skips e.g. "vertical-align")
TEXT_ALIGN_REGEX = re.compile(r'(?<![\w-])text-align\s*:\s*(\w+)', re.IGNORECASE)

# Background values which mean "no background"
NO_BACKGROUND_VALUES = {'transparent', 'none', 'initial', 'inherit'}


def is_color_span(tag):
    """True for the <span style="color:..."> elements created by this app."""
    return (tag.name == 'span'
            and (tag.get('style') or '').startswith('color:'))


def is_background_span(tag):
    """True for the <span style="background-color:..."> elements created by this app."""
    return (tag.name == 'span'
            and (tag.get('style') or '').startswith('background-color:'))


def is_deco_span(tag):
    """True for any span created by this app (text colour or background)."""
    return is_color_span(tag) or is_background_span(tag)


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


def make_check_mark_file(color="#ffffff", size=18):
    """Draws a check mark (inline SVG) into a PNG in the temp folder and returns
    its path with forward slashes, ready for a stylesheet url(...). Qt style
    sheets need a file for QCheckBox::indicator:checked images."""
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 18 18">'
        f'<path d="M4 9.5 L7.5 13 L14 5.5" fill="none" stroke="{color}" '
        'stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"/>'
        '</svg>'
    )
    renderer = QSvgRenderer(QByteArray(svg.encode("utf-8")))
    image = QImage(size * 2, size * 2, QImage.Format.Format_ARGB32)
    image.fill(Qt.GlobalColor.transparent)
    painter = QPainter(image)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    renderer.render(painter)
    painter.end()
    path = os.path.join(tempfile.gettempdir(), "links2html_check.png")
    image.save(path)
    return path.replace("\\", "/")


class SettingsDialog(QDialog):
    """Small settings window with the app options as checkboxes
    (built from SETTINGS_OPTIONS; self.checkboxes maps key -> QCheckBox)."""

    def __init__(self, parent, values):
        super().__init__(parent)
        self.setWindowTitle("Settings")
        self.setModal(True)
        self.setMinimumWidth(460)

        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        self.checkboxes = {}
        for key, label, _default in SETTINGS_OPTIONS:
            cb = QCheckBox(label)
            cb.setChecked(values[key])
            self.checkboxes[key] = cb
            layout.addWidget(cb)
        layout.addStretch()

        btn_close = QPushButton("Close")
        btn_close.clicked.connect(self.accept)
        layout.addWidget(btn_close)


class DocumentToHtmlConverter(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Link Converter: Word & PowerPoint -> HTML (Dark Theme)")
        self.setWindowIcon(load_app_icon())
        self.resize(1000, 600)

        # --- Settings (persisted via QSettings; defaults in SETTINGS_OPTIONS) ---
        # type=bool is required: QSettings may return "true"/"false" strings
        # (e.g. from an .ini backend), which would otherwise both be truthy.
        self.qsettings = QSettings("Links2HTML", "Links2HTML")
        self.opts = {key: self.qsettings.value(key, default, type=bool)
                     for key, _label, default in SETTINGS_OPTIONS}
 
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
        self.preview_area.setObjectName("previewArea")
        self.preview_area.textChanged.connect(self.on_preview_changed)
 
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
            QTextEdit#previewArea {
                color: %s;
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
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border: 1px solid #666666;
                border-radius: 4px;
                background-color: #1e1e1e;
            }
            QCheckBox::indicator:hover {
                border-color: #436a9b;
            }
            QCheckBox::indicator:checked {
                background-color: #365880;
                border-color: #436a9b;
                image: url("%s");
            }
            QCheckBox::indicator:checked:hover {
                background-color: #436a9b;
            }
        """
        self.setStyleSheet(dark_stylesheet % (PREVIEW_TEXT_COLOR,
                                              make_check_mark_file()))

    # ------------------------------------------------------------------
    # Settings
    # ------------------------------------------------------------------
    def open_settings(self):
        dlg = SettingsDialog(self, self.opts)
        dlg.setWindowIcon(self.windowIcon())
        for key, cb in dlg.checkboxes.items():
            cb.toggled.connect(partial(self.set_option, key))
        dlg.exec()

    def set_option(self, key, checked):
        self.opts[key] = checked
        self.qsettings.setValue(key, checked)
        self.process_content()  # refresh the result immediately

    def on_preview_changed(self):
        self.force_preview_text_color()
        self.process_content()

    def force_preview_text_color(self):
        """Displays ALL text of the preview window in grey, whatever colour the
        pasted/loaded content has. Done with an "extra selection" (a paint-time
        overlay), so the document itself - and therefore the colour detection
        in process_content() - is left untouched."""
        cursor = QTextCursor(self.preview_area.document())
        cursor.select(QTextCursor.SelectionType.Document)
        selection = QTextEdit.ExtraSelection()
        selection.cursor = cursor
        selection.format.setForeground(QColor(PREVIEW_TEXT_COLOR))
        self.preview_area.setExtraSelections([selection])
 
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
                        "strike => s\n"
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
                        rPr = run._r.find(qn('a:rPr'))

                        # Text highlight (<a:highlight><a:srgbClr val="..."/></a:highlight>)
                        # -> background-color span. Written into the preview only;
                        # process_content() decides whether it is kept (Settings).
                        if not is_link and rPr is not None:
                            highlight = rPr.find(qn('a:highlight'))
                            if highlight is not None:
                                srgb = highlight.find(qn('a:srgbClr'))
                                if srgb is not None and srgb.get('val'):
                                    text = f'<span style="background-color:#{srgb.get("val")}">{text}</span>'

                        if not is_link:
                            try:
                                if font.color and font.color.type == MSO_COLOR_TYPE.RGB:
                                    text = f'<span style="color:#{font.color.rgb}">{text}</span>'
                            except AttributeError:
                                pass

                        # python-pptx has no direct superscript/subscript API,
                        # so read the raw <a:rPr baseline="..."/> attribute:
                        # positive baseline => superscript, negative => subscript.
                        baseline = rPr.get('baseline') if rPr is not None else None
                        is_superscript = baseline is not None and int(baseline) > 0
                        is_subscript = baseline is not None and int(baseline) < 0

                        if is_superscript:
                            text = f"<sup>{text}</sup>"
                        if is_subscript:
                            text = f"<sub>{text}</sub>"
                        if font.underline and not is_link:
                            text = f"<u>{text}</u>"
                        if rPr is not None and rPr.get('strike') in ('sngStrike', 'dblStrike'):
                            text = f"<s>{text}</s>"
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
                        align = {PP_ALIGN.CENTER: "center", PP_ALIGN.RIGHT: "right",
                                 PP_ALIGN.JUSTIFY: "justify"}.get(paragraph.alignment)
                        style_attr = f' style="text-align:{align}"' if align else ""
                        html_output += f"<p{style_attr}>{p_html}</p>\n"

        return html_output
 
    def paste_from_clipboard(self):
        self.preview_area.clear()
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

    @staticmethod
    def clean_inline(node, keep_tags):
        """Unwraps every tag of the node except the allowed inline ones (and our
        colour/background spans); returns the resulting inner HTML."""
        for tag in [t for t in node.find_all(True)
                    if t.name not in keep_tags and not is_deco_span(t)]:
            tag.unwrap()
        return "".join(str(c) for c in node.contents).strip()

    def render_list(self, list_tag, keep_tags):
        """Renders a <ul>/<ol> (with nested lists) as clean HTML.
        Qt puts nested lists directly inside the parent list, whereas
        mammoth/HTML puts them inside an <li> - both forms are handled."""
        items = []
        for child in list_tag.children:
            if getattr(child, 'name', None) == 'li':
                nested = "".join(self.render_list(n, keep_tags)
                                 for n in child.find_all(['ul', 'ol'], recursive=False))
                for n in child.find_all(['ul', 'ol'], recursive=False):
                    n.extract()
                items.append(f"<li>{self.clean_inline(child, keep_tags)}{nested}</li>")
            elif getattr(child, 'name', None) in ('ul', 'ol'):
                nested = self.render_list(child, keep_tags)
                if items:
                    items[-1] = items[-1][:-len("</li>")] + nested + "</li>"
                else:
                    items.append(f"<li>{nested}</li>")
        if not items:
            return ""
        return f"<{list_tag.name}>\n" + "\n".join(items) + f"\n</{list_tag.name}>"

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
            deco_match = TEXT_DECORATION_REGEX.search(style)
            decoration = deco_match.group(1).lower() if deco_match else ''
            is_underline = 'underline' in decoration
            is_strike = self.opts["detect_strikethrough"] and 'line-through' in decoration
            is_superscript = ('vertical-align:super' in style or
                              'vertical-align: super' in style)
            is_subscript = ('vertical-align:sub' in style or
                            'vertical-align: sub' in style)

            # Text colour (optional feature). Spans that sit inside a link are
            # ignored, so links never get a colour of their own.
            color_value = None
            if self.opts["detect_text_color"] and not span.find_parent('a'):
                color_match = COLOR_STYLE_REGEX.search(style)
                if color_match:
                    color_value = color_match.group(1).strip()

            # Text background / highlight (optional feature), same rules as colour.
            background_value = None
            if self.opts["detect_background"] and not span.find_parent('a'):
                bg_match = BACKGROUND_STYLE_REGEX.search(style)
                if bg_match:
                    candidate = bg_match.group(1).strip()
                    if candidate.lower() not in NO_BACKGROUND_VALUES:
                        background_value = candidate

            if not (is_bold or is_italic or is_underline or is_strike or
                    is_superscript or is_subscript or color_value or background_value):
                continue

            # Collect the span's children, then wrap them in layers:
            # background → color → sup/sub → u → s → i → b
            # (innermost first so the outermost tag is the first one readers see)
            children = list(span.contents)

            if background_value:
                bg_tag = soup.new_tag('span')
                bg_tag['style'] = f"background-color:{background_value}"
                for child in children:
                    bg_tag.append(child)
                children = [bg_tag]

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

            if is_strike:
                s_tag = soup.new_tag('s')
                for child in children:
                    s_tag.append(child)
                children = [s_tag]

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
        for old_strike in body.find_all(['strike', 'del']):
            old_strike.name = 's'

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
        # The same is done for background spans.
        if self.opts["detect_text_color"] or self.opts["detect_background"]:
            for deco_span in body.find_all(is_deco_span):
                if deco_span.find('a') is None:
                    continue
                deco_style = deco_span['style']
                for text_node in deco_span.find_all(string=True):
                    if text_node.find_parent('a') or not text_node.strip():
                        continue
                    wrapper = soup.new_tag('span')
                    wrapper['style'] = deco_style
                    text_node.replace_with(wrapper)
                    wrapper.append(text_node)
                deco_span.unwrap()

        # --- Step 2e: Open links in a new tab (optional feature) ---
        if self.opts["links_new_tab"]:
            for a_tag in body.find_all('a'):
                a_tag['target'] = '_blank'

        # --- Step 3: Build clean output — keep <a>, <b>, <i>, <u>, <sub>, <sup> ---
        # (plus colour/background spans created above, and <s>, headings,
        # lists and paragraph alignment when the matching options are enabled)
        keep_tags = {'a', 'b', 'i', 'u', 'sub', 'sup'}
        if self.opts["detect_strikethrough"]:
            keep_tags.add('s')
        keep_lists = self.opts["keep_lists"]
        keep_headings = self.opts["keep_headings"]

        block_tags = ['p', 'div'] + HEADING_TAGS
        block_tags += ['ul', 'ol'] if keep_lists else ['li']

        parts = []
        for block in body.find_all(block_tags):
            # Already consumed (extracted/unwrapped while rendering its parent)
            if block.parent is None:
                continue

            # Lists: a top-level <ul>/<ol> is rendered as a whole (nested lists
            # included); every block sitting inside a list is already part of it.
            if keep_lists:
                if block.name in ('ul', 'ol'):
                    if not block.find_parent(['ul', 'ol']):
                        list_html = self.render_list(block, keep_tags)
                        if list_html:
                            parts.append(list_html)
                    continue
                if block.find_parent(['ul', 'ol']):
                    continue

            # Empty line (Qt: <p><br></p>) -> <br><br> (optional feature).
            # Only leaf paragraphs count, so wrappers around real content
            # never produce a stray <br><br>.
            if (self.opts["empty_line_br"] and block.name in ('p', 'div')
                    and not block.get_text(strip=True)
                    and not block.find(block_tags)):
                parts.append(EMPTY_LINE_HTML)
                continue

            is_heading = block.name in HEADING_TAGS
            keeps_tag = is_heading and keep_headings
            styles = []

            # If the WHOLE paragraph has one colour, the colour is moved from
            # the <span>s to the paragraph itself: <p style="color:#...">...</p>.
            # Mixed/partial colouring keeps inline <span>s (a <p> can't be
            # placed in the middle of a sentence).
            if self.opts["detect_text_color"] and (block.name in ('p', 'div') or keeps_tag):
                uniform_style = self.uniform_block_color(block)
                if uniform_style:
                    styles.append(uniform_style)
                    for color_span in block.find_all(is_color_span):
                        color_span.unwrap()

            # Paragraph alignment (optional feature); left is the default, skipped.
            if self.opts["keep_alignment"] and (block.name in ('p', 'div') or keeps_tag):
                # Qt writes it as align="..." attribute, HTML sources as CSS.
                align_match = TEXT_ALIGN_REGEX.search(block.get('style') or '')
                align = (align_match.group(1) if align_match
                         else block.get('align') or '').strip().lower()
                if align in ('center', 'right', 'justify'):
                    styles.append(f"text-align:{align}")

            # Headings are bold anyway - drop the redundant <b> inside them.
            if keeps_tag:
                for b_tag in block.find_all('b'):
                    b_tag.unwrap()

            block_html = self.clean_inline(block, keep_tags)
            if not block_html:
                continue

            tag_name = block.name if keeps_tag else ('p' if styles else None)
            if tag_name:
                style_attr = (f' style="{html.escape(";".join(styles), quote=True)}"'
                              if styles else '')
                block_html = f'<{tag_name}{style_attr}>{block_html}</{tag_name}>'
            parts.append(block_html)

        # Empty lines at the very start / end of the text are not content
        while parts and parts[0] == EMPTY_LINE_HTML:
            parts.pop(0)
        while parts and parts[-1] == EMPTY_LINE_HTML:
            parts.pop()

        self.result_area.setPlainText("\n\n".join(parts))
 
if __name__ == "__main__":
    # Windows: own AppUserModelID, so the taskbar shows our icon instead of python's
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Links2HTML.Links2HTML")
        except Exception:
            pass

    app = QApplication(sys.argv)
    app.setWindowIcon(load_app_icon())  # main window, dialogs and message boxes
    window = DocumentToHtmlConverter()
    window.show()
    sys.exit(app.exec())
