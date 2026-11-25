from pathlib import Path
import json
import markdown
import bleach
from PySide6.QtCore import Signal, QUrl, QTimer
from PySide6.QtGui import QDesktopServices
from PySide6.QtWebEngineCore import QWebEnginePage
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWidgets import QMenu
from lang import t


MARKDOWN_TAGS = set(bleach.sanitizer.ALLOWED_TAGS) | {
    "h1", "h2", "h3", "h4", "h5", "h6", "p", "br", "hr", "pre",
    "div", "span", "img", "table", "thead", "tbody", "tr", "th", "td",
}
MARKDOWN_ATTRIBUTES = {
    "a": ["href", "title", "name", "id"],
    "h1": ["id"], "h2": ["id"], "h3": ["id"],
    "h4": ["id"], "h5": ["id"], "h6": ["id"],
    "img": ["src", "alt", "title", "width", "height"],
    "code": ["class"],
    "pre": ["class"],
}


def sanitize_markdown_html(html_body):
    return bleach.clean(
        html_body,
        tags=MARKDOWN_TAGS,
        attributes=MARKDOWN_ATTRIBUTES,
        strip=True,
    )


def is_internal_anchor_url(url, current_file_path):
    if not url.fragment() or url.scheme() not in ("", "file"):
        return False
    if not url.path():
        return True

    current_path = Path(current_file_path).resolve()
    linked_path = (
        Path(url.toLocalFile()).resolve()
        if url.isLocalFile()
        else Path(url.path()).resolve()
    )
    return linked_path in (current_path, current_path.parent)


class CustomMarkdownPage(QWebEnginePage):
    link_clicked = Signal(QUrl)

    def acceptNavigationRequest(self, url, nav_type, is_main_frame):
        if nav_type == QWebEnginePage.NavigationTypeLinkClicked:
            self.link_clicked.emit(url)
            return False
        return super().acceptNavigationRequest(url, nav_type, is_main_frame)


class MarkdownDocAdapter:
    """MainWindow.load_pdf_bookmarks() ile tam uyum sağlayan içindekiler adaptörü."""
    def __init__(self, toc_list):
        self._toc = toc_list  # [(level, title, anchor_id), ...]

    def get_toc(self):
        return self._toc

    def close(self):
        pass

    def __len__(self):
        return 1


class MarkdownViewer(QWebEngineView):
    zoom_changed = Signal()
    page_changed = Signal(int)
    linkClicked = Signal()
    bookmarkRequested = Signal(int, str, float, float)

    def __init__(self, theme="light", parent=None):
        super().__init__(parent)
        self.doc = None
        self.current_page = 0
        self.zoom = 1.0
        self.current_file_path = ""
        self.original_file_path = ""
        self.theme = theme
        self._is_loaded = False
        self._pending_scroll_y = None
        self._pending_anchor = None
        self._pending_occurrence = None

        self.current_scroll_y = 0.0
        self.page().scrollPositionChanged.connect(self._on_scroll_changed)
        self.titleChanged.connect(self._on_title_changed)

        self.custom_page = CustomMarkdownPage(self)
        self.custom_page.link_clicked.connect(self._handle_link)
        self.setPage(self.custom_page)
        self.loadFinished.connect(self._on_load_finished)

    def _on_scroll_changed(self, pos):
        self.current_scroll_y = float(pos.y())

    def _on_title_changed(self, title):
        if title.startswith("SCROLL:"):
            try:
                self.current_scroll_y = float(title.split(":", 1)[1])
            except ValueError:
                pass

    def _on_load_finished(self, ok):
        self._is_loaded = True
        if not ok:
            return

        if self._pending_anchor:
            anchor, title = self._pending_anchor
            self._pending_anchor = None
            self.scroll_to_anchor(anchor, title)
        elif self._pending_scroll_y is not None:
            y_val = int(self._pending_scroll_y)
            self._pending_scroll_y = None
            self._apply_scroll_y(y_val)
        if self._pending_occurrence is not None:
            query, occurrence_index = self._pending_occurrence
            self._pending_occurrence = None
            self.jump_to_occurrence(query, occurrence_index)

    def _apply_scroll_y(self, y_val):
        self.current_scroll_y = float(y_val)
        js = f"window.scrollTo({{top: {y_val}, left: 0, behavior: 'instant'}});"
        self.page().runJavaScript(js, lambda _: self._force_repaint())
        QTimer.singleShot(120, lambda: self.page().runJavaScript(js, lambda _: self._force_repaint()))

    def scroll_to_anchor(self, anchor_id, fallback_title=""):
        if not self._is_loaded:
            self._pending_anchor = (anchor_id, fallback_title)
            return

        safe_id = json.dumps(anchor_id)
        safe_title = json.dumps(fallback_title.strip())

        js = f"""
        (function() {{
            var targetId = {safe_id};
            var targetTitle = {safe_title};
            var el = document.getElementById(targetId);

            if (!el && targetTitle) {{
                var headings = document.querySelectorAll('h1, h2, h3, h4, h5, h6');
                for (var i = 0; i < headings.length; i++) {{
                    if (headings[i].textContent.trim() === targetTitle) {{
                        el = headings[i];
                        break;
                    }}
                }}
            }}

            if (el) {{
                el.scrollIntoView({{behavior: 'instant', block: 'start'}});
            }}
            return window.scrollY || 0;
        }})();
        """
        def _after_scroll(new_y):
            if new_y is not None:
                self.current_scroll_y = float(new_y)
            self._force_repaint()

        self.page().runJavaScript(js, _after_scroll)

    def _force_repaint(self):
        self.setFocus()
        self.update()
        self.page().runJavaScript("requestAnimationFrame(() => {});")

    def open_markdown(self, file_path):
        self._is_loaded = False
        self.current_file_path = str(file_path)
        self.original_file_path = str(file_path)
        p = Path(file_path)

        md_text = p.read_text(encoding="utf-8")
        md_converter = markdown.Markdown(
            extensions=["toc", "tables", "fenced_code", "sane_lists", "nl2br"]
        )
        html_body = sanitize_markdown_html(md_converter.convert(md_text))

        toc_items = []
        def _flatten_toc(tokens):
            for tok in tokens:
                toc_items.append((tok["level"], tok["name"], tok["id"]))
                if tok.get("children"):
                    _flatten_toc(tok["children"])

        _flatten_toc(getattr(md_converter, "toc_tokens", []))
        self.doc = MarkdownDocAdapter(toc_items)

        is_dark = self.theme == "dark"
        bg_color = "#151a21" if is_dark else "#ffffff"
        text_color = "#d1d6dc" if is_dark else "#24292f"
        border_color = "#2d3640" if is_dark else "#d0d7de"
        code_bg = "#1e2630" if is_dark else "#f6f8fa"
        link_color = "#58a6ff" if is_dark else "#0969da"

        full_html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <style>
                body {{
                    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif;
                    background-color: {bg_color};
                    color: {text_color};
                    line-height: 1.6;
                    max-width: 960px;
                    margin: 0 auto;
                    padding: 24px 32px;
                    word-wrap: break-word;
                }}
                img {{ max-width: 100%; height: auto; border-radius: 6px; }}
                h1, h2, h3, h4 {{
                    border-bottom: 1px solid {border_color};
                    padding-bottom: 0.3em;
                    margin-top: 24px;
                }}
                a {{ color: {link_color}; text-decoration: none; }}
                a:hover {{ text-decoration: underline; }}
                code {{
                    background-color: {code_bg};
                    padding: 0.2em 0.4em;
                    border-radius: 6px;
                    font-family: monospace;
                }}
                pre {{
                    background-color: {code_bg};
                    padding: 16px;
                    border-radius: 6px;
                    overflow-x: auto;
                }}
                pre code {{ background: none; padding: 0; }}
                table {{ border-collapse: collapse; width: 100%; margin: 16px 0; }}
                th, td {{ border: 1px solid {border_color}; padding: 8px 12px; }}
                blockquote {{
                    border-left: 4px solid {border_color};
                    margin: 0;
                    padding-left: 16px;
                    opacity: 0.85;
                }}
            </style>
        </head>
        <body>
            {html_body}
            <script>
                window.addEventListener('scroll', function() {{
                    document.title = "SCROLL:" + Math.round(window.scrollY);
                }}, {{passive: true}});
            </script>
        </body>
        </html>
        """
        base_url = QUrl.fromLocalFile(str(p.parent) + "/")
        self.setHtml(full_html, base_url)

    def contextMenuEvent(self, event):
        menu = QMenu(self)
        copy_action = menu.addAction(t("view_copy"))
        bookmark_action = menu.addAction(t("view_add_bookmark"))

        selected_text = self.selectedText().strip()
        copy_action.setEnabled(bool(selected_text))

        action = menu.exec(event.globalPos())
        if action == copy_action:
            self.triggerPageAction(QWebEnginePage.Copy)
        elif action == bookmark_action:
            self._request_bookmark_from_selection(selected_text)

    def _request_bookmark_from_selection(self, selected_text):
        title = " ".join(selected_text.split()) if selected_text else Path(self.current_file_path).stem
        if len(title) > 60:
            title = title[:60] + "..."

        self.page().runJavaScript(
            "window.pageYOffset || document.documentElement.scrollTop || 0",
            lambda scroll_y: self.bookmarkRequested.emit(0, title, 0.0, float(scroll_y or 0.0))
        )

    def _handle_link(self, url: QUrl):
        self.linkClicked.emit()
        local_file = url.toLocalFile()
        fragment = url.fragment()

        if is_internal_anchor_url(url, self.current_file_path):
            self.scroll_to_anchor(fragment)
            return

        if url.isLocalFile() and Path(local_file).exists():
            main_win = self.window()
            if hasattr(main_win, "open_in_tab"):
                main_win.open_in_tab(local_file)
                if fragment:
                    new_v = main_win.get_current_viewer()
                    if isinstance(new_v, MarkdownViewer):
                        new_v.scroll_to_anchor(fragment)
            return

        QDesktopServices.openUrl(url)

    def goto_position(self, page, x, y):
        y_val = int(y or 0)
        if self._is_loaded:
            self._apply_scroll_y(y_val)
        else:
            self._pending_scroll_y = y_val

    def goto_page(self, page, title=""):
        if isinstance(page, str):
            self.scroll_to_anchor(page, title)

    def next_page(self):
        self.page().runJavaScript("window.scrollBy(0, window.innerHeight * 0.8);")

    def previous_page(self):
        self.page().runJavaScript("window.scrollBy(0, -window.innerHeight * 0.8);")

    def zoom_in(self):
        self.zoom = min(3.0, self.zoom + 0.1)
        self.setZoomFactor(self.zoom)

    def zoom_out(self):
        self.zoom = max(0.5, self.zoom - 0.1)
        self.setZoomFactor(self.zoom)

    def zoom_100(self):
        self.zoom = 1.0
        self.setZoomFactor(1.0)

    def fit_to_width(self):
        self.zoom_100()

    def jump_to_occurrence(self, query_str, occurrence_index=0):
        """Sayfadaki N. (0-tabanlı) kelime eşleşmesine doğrudan kaydırır ve vurgular."""
        if not query_str:
            return

        safe_q = json.dumps(query_str)
        idx = int(occurrence_index)

        js = f"""
        (function() {{
            // Önceki özel arama vurgularını temizle
            document.querySelectorAll('mark.docnet-search-hit').forEach(function(m) {{
                var parent = m.parentNode;
                parent.replaceChild(document.createTextNode(m.textContent), m);
                parent.normalize();
            }});

            var query = {safe_q}.toLowerCase();
            var targetIdx = {idx};
            var currentIdx = 0;
            var targetMark = null;

            var walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null, false);
            var textNodes = [];
            var node;
            while ((node = walker.nextNode())) {{
                if (node.parentElement && ['SCRIPT', 'STYLE'].indexOf(node.parentElement.tagName) === -1) {{
                    if (node.nodeValue.toLowerCase().indexOf(query) !== -1) {{
                        textNodes.push(node);
                    }}
                }}
            }}

            for (var i = 0; i < textNodes.length; i++) {{
                var tNode = textNodes[i];
                var text = tNode.nodeValue;
                var lowerText = text.toLowerCase();
                var startPos = 0;
                var matchPos;

                var frag = document.createDocumentFragment();
                var modified = false;

                while ((matchPos = lowerText.indexOf(query, startPos)) !== -1) {{
                    if (matchPos > startPos) {{
                        frag.appendChild(document.createTextNode(text.slice(startPos, matchPos)));
                    }}
                    var mark = document.createElement('mark');
                    mark.className = 'docnet-search-hit';
                    mark.textContent = text.slice(matchPos, matchPos + query.length);

                    if (currentIdx === targetIdx) {{
                        mark.style.backgroundColor = '#ff9800';
                        mark.style.color = '#000000';
                        mark.style.padding = '1px 3px';
                        mark.style.borderRadius = '3px';
                        targetMark = mark;
                    }} else {{
                        mark.style.backgroundColor = '#ffeb3b';
                        mark.style.color = '#000000';
                    }}

                    frag.appendChild(mark);
                    currentIdx++;
                    startPos = matchPos + query.length;
                    modified = true;
                }}

                if (modified) {{
                    if (startPos < text.length) {{
                        frag.appendChild(document.createTextNode(text.slice(startPos)));
                    }}
                    tNode.parentNode.replaceChild(frag, tNode);
                }}
            }}

            if (targetMark) {{
                targetMark.scrollIntoView({{behavior: 'instant', block: 'center'}});
            }}
            return window.scrollY || 0;
        }})();
        """
        def _after_jump(new_y):
            if new_y is not None:
                self.current_scroll_y = float(new_y)
            self._force_repaint()

        if self._is_loaded:
            self.page().runJavaScript(js, _after_jump)
        else:
            self._pending_occurrence = (query_str, idx)
