import pymupdf
import pytesseract
from PIL import Image

from PySide6.QtCore import Signal, Qt, QRect, QUrl, QMimeData, QThread, QTimer
from PySide6.QtGui import (
    QPixmap,
    QImage,
    QPainter,
    QColor,
    QPen,
    QShortcut,
    QKeySequence,
    QDesktopServices,
)
from PySide6.QtWidgets import (
    QLabel,
    QScrollArea,
    QApplication,
    QMenu,
    QRubberBand,
)
from lang import t


MAX_ZOOM = 4.0


def screen_to_pdf_point(screen_point, page, zoom):
    matrix = page.rotation_matrix * pymupdf.Matrix(zoom, zoom)
    crop_rect = page.cropbox * matrix
    rendered_point = pymupdf.Point(screen_point.x(), screen_point.y())
    return (rendered_point + crop_rect.top_left) * ~matrix


def pdf_to_screen_point(pdf_point, page, zoom):
    matrix = page.rotation_matrix * pymupdf.Matrix(zoom, zoom)
    crop_rect = page.cropbox * matrix
    return pdf_point * matrix - crop_rect.top_left


def pdf_to_screen_rect(pdf_rect, page, zoom):
    points = [
        pdf_to_screen_point(pymupdf.Point(x, y), page, zoom)
        for x, y in (
            (pdf_rect.x0, pdf_rect.y0),
            (pdf_rect.x1, pdf_rect.y0),
            (pdf_rect.x1, pdf_rect.y1),
            (pdf_rect.x0, pdf_rect.y1),
        )
    ]
    return pymupdf.Rect(
        min(point.x for point in points),
        min(point.y for point in points),
        max(point.x for point in points),
        max(point.y for point in points),
    )


class PageRenderWorker(QThread):
    rendered = Signal(int, int, float, QImage)
    error = Signal(int, str)

    def __init__(self, request_id, file_path, page_no, zoom):
        super().__init__()
        self.request_id = request_id
        self.file_path = file_path
        self.page_no = page_no
        self.zoom = zoom

    def run(self):
        try:
            if self.isInterruptionRequested():
                return
            with pymupdf.open(self.file_path) as doc:
                pix = doc[self.page_no].get_pixmap(
                    matrix=pymupdf.Matrix(self.zoom, self.zoom),
                    annots=True,
                )
                if self.isInterruptionRequested():
                    return
                image = QImage(
                    pix.samples,
                    pix.width,
                    pix.height,
                    pix.stride,
                    QImage.Format_RGB888,
                ).copy()
            self.rendered.emit(
                self.request_id, self.page_no, self.zoom, image
            )
        except Exception as e:
            self.error.emit(self.request_id, str(e))


class OCRThread(QThread):
    recognized = Signal(str)
    error = Signal(str)

    def __init__(self, img, ocr_lang: str = "tur+eng") -> None:
        super().__init__()
        self.img = img
        self.ocr_lang = ocr_lang

    def run(self) -> None:
        try:
            # Hardcoded değer yerine değişkenden okuyoruz
            text = pytesseract.image_to_string(self.img, lang=self.ocr_lang, config="--psm 6")
            self.recognized.emit(text.strip())
        except Exception as e:
            self.error.emit(str(e))

class DOCViewer(QScrollArea):

    zoom_changed = Signal()
    bookmarkRequested = Signal(int, str, float, float)
    highlightRequested = Signal(int, float, float, float, float)
    deleteHighlightRequested = Signal(int, float, float)
    page_changed = Signal(int)
    exportTableRequested = Signal(int, float, float, float, float)
    addNoteRequested = Signal(int, float, float)
    linkClicked = Signal()
    renderError = Signal(str)


    def __init__(self):
        super().__init__()
        self.doc = None
        self.current_page = 0
        self.zoom = 1.5
        self.current_file_path = ""
        self._render_request_id = 0
        self._render_worker = None
        self._render_pending = False
        self._pending_position = None
        self._pending_scroll_top = False

        self.render_timer = QTimer(self)
        self.render_timer.setSingleShot(True)
        self.render_timer.timeout.connect(self._start_page_render)
        self.image_label = QLabel()
        self.image_label.setAlignment(Qt.AlignCenter)
        self.image_label.mousePressEvent = self.mousePressEvent
        self.image_label.mouseMoveEvent = self.mouseMoveEvent
        self.image_label.mouseReleaseEvent = self.mouseReleaseEvent

        self.image_label.setMouseTracking(True)
        self.setMouseTracking(True)

        self.setWidget(self.image_label)
        self.setWidgetResizable(False)
        self.setAlignment(Qt.AlignCenter)

        self.rubberBand = QRubberBand(QRubberBand.Rectangle, self.image_label)

        self.highlight_rects = []
        self.selecting = False
        self.selection_start = None
        self.selection_end = None
        self.selection_page = None

        self.pan_mode = False
        self.pan_start_pos = None
        self.dragged_enough = False
        self.active_search_rect = None
        self.ocr_thread = None

        self.copy_shortcut = QShortcut(QKeySequence.Copy, self)
        self.copy_shortcut.activated.connect(self.copy_selected_text)

        self.page_up_shortcut = QShortcut(QKeySequence(Qt.Key_PageUp), self)
        self.page_up_shortcut.setContext(Qt.WidgetWithChildrenShortcut)
        self.page_up_shortcut.activated.connect(self.previous_page)

        self.page_down_shortcut = QShortcut(QKeySequence(Qt.Key_PageDown), self)
        self.page_down_shortcut.setContext(Qt.WidgetWithChildrenShortcut)
        self.page_down_shortcut.activated.connect(self.next_page)

    def open_pdf(self, pdf_path):
        self._render_request_id += 1
        self.render_timer.stop()
        self._pending_position = None
        self._pending_scroll_top = False
        if self._render_worker and self._render_worker.isRunning():
            self._render_worker.requestInterruption()
        if self.doc:
            self.doc.close()
        self.current_file_path = pdf_path
        try:
            self.doc = pymupdf.open(pdf_path)
            self.current_page = 0
            self.highlight_rects = []
            self.render_page()
        except Exception:
            from PySide6.QtWidgets import QMessageBox
            QMessageBox.critical(self, t("err_open_title"), t("err_open_msg"))
            self.doc = None

    def _start_page_render(self):
        if not self.doc or not self.current_file_path:
            return
        if self._render_worker and self._render_worker.isRunning():
            self._render_pending = True
            self._render_worker.requestInterruption()
            return
        if self._render_worker:
            self._render_worker.wait()
            self._render_worker.deleteLater()

        self._render_pending = False
        worker = PageRenderWorker(
            self._render_request_id,
            self.current_file_path,
            self.current_page,
            self.zoom,
        )
        worker.rendered.connect(self._on_page_rendered)
        worker.error.connect(self._on_page_render_error)
        worker.finished.connect(
            lambda worker=worker: self._on_render_worker_finished(worker)
        )
        self._render_worker = worker
        worker.start()

    def _on_page_rendered(self, request_id, page_no, zoom, image):
        if (
            request_id != self._render_request_id
            or not self.doc
            or page_no != self.current_page
            or zoom != self.zoom
        ):
            return

        pixmap = QPixmap.fromImage(image)
        painter = QPainter(pixmap)
        page = self.doc[page_no]

        for highlight_page, rect in self.highlight_rects:
            if highlight_page != page_no:
                continue
            screen_rect = pdf_to_screen_rect(rect, page, zoom)
            overlay = QRect(
                round(screen_rect.x0),
                round(screen_rect.y0),
                round(screen_rect.width),
                round(screen_rect.height),
            )
            is_active = (
                self.active_search_rect is not None
                and abs(rect.x0 - self.active_search_rect.x0) < 0.1
                and abs(rect.y0 - self.active_search_rect.y0) < 0.1
            )
            if is_active:
                painter.fillRect(overlay, QColor(0, 255, 150, 120))
                painter.setPen(QPen(QColor(0, 50, 255), 3))
            else:
                painter.fillRect(overlay, QColor(255, 255, 0, 80))
                painter.setPen(QPen(QColor(255, 0, 0), 1.5))
            painter.drawRect(overlay)

        if (
            self.selection_page == page_no
            and self.selection_start is not None
            and self.selection_end is not None
        ):
            selection = QRect(self.selection_start, self.selection_end).normalized()
            painter.fillRect(selection, QColor(0, 120, 255, 60))
            painter.setPen(QPen(QColor(0, 120, 255), 2))
            painter.drawRect(selection)

        painter.end()
        self.image_label.setPixmap(pixmap)
        self.image_label.resize(pixmap.size())
        if self._pending_position is not None:
            x, y = self._pending_position
            self._pending_position = None
            point = pdf_to_screen_point(
                pymupdf.Point(x, y), page, zoom
            )
            self.horizontalScrollBar().setValue(round(point.x))
            self.verticalScrollBar().setValue(round(point.y))
        elif self._pending_scroll_top:
            self._pending_scroll_top = False
            self.horizontalScrollBar().setValue(0)
            self.verticalScrollBar().setValue(0)

    def _on_page_render_error(self, request_id, message):
        if request_id == self._render_request_id:
            self.renderError.emit(message)

    def _on_render_worker_finished(self, worker):
        if self._render_worker is worker:
            self._render_worker = None
            worker.deleteLater()
            if self._render_pending:
                self._start_page_render()

    def wait_for_render(self):
        self.render_timer.stop()
        self._render_request_id += 1
        self._render_pending = False
        worker = self._render_worker
        if worker:
            if worker.isRunning():
                worker.requestInterruption()
                worker.wait()
            self._render_worker = None
            worker.deleteLater()

    # Yeni render_page fonksiyonumuz sadece timer'ı tetikliyor
    def render_page(self) -> None:
        self._render_request_id += 1
        self.render_timer.start(150)

    def set_highlight(self, rects):
        self.highlight_rects = rects
        self.render_page()

    def next_page(self):
        if self.doc and self.current_page < len(self.doc) - 1:
            self.highlight_rects = []
            self.active_search_rect = None
            self.current_page += 1
            self.render_page()
            self.page_changed.emit(self.current_page)

    def previous_page(self):
        if self.doc and self.current_page > 0:
            self.highlight_rects = []
            self.active_search_rect = None
            self.current_page -= 1
            self.render_page()
            self.page_changed.emit(self.current_page)

    def goto_page(self, page):
        if not self.doc or page < 0 or page >= len(self.doc):
            return
        self.selection_start = None
        self.selection_end = None
        self.highlight_rects = []
        self.active_search_rect = None

        self._pending_position = None
        self._pending_scroll_top = False
        self.current_page = page
        self.render_page()
        self.page_changed.emit(self.current_page)

    def goto_position(self, page, x, y):
        if not self.doc or page < 0 or page >= len(self.doc):
            return
        self.goto_page(page)
        self._pending_position = (x, y)

    def goto_page_top(self, page):
        if not self.doc or page < 0 or page >= len(self.doc):
            return
        self.goto_page(page)
        self._pending_scroll_top = True

    def get_pdf_position_from_scroll(self):
        if not self.doc:
            return 0.0, 0.0
        viewport_origin = self.viewport().mapTo(
            self.image_label, self.viewport().rect().topLeft()
        )
        point = screen_to_pdf_point(
            viewport_origin,
            self.doc[self.current_page],
            self.zoom,
        )
        return point.x, point.y

    def search_text(self, text):
        if not self.doc: return []
        results = []
        for page_no in range(len(self.doc)):
            page = self.doc[page_no]
            for rect in page.search_for(text):
                results.append((page_no, rect))
        return results

    def get_pdf_point(self, event):
        local_pos = self.image_label.mapFromGlobal(event.globalPosition().toPoint())
        if local_pos.x() < 0 or local_pos.y() < 0:
            return None
        if not self.doc:
            return None
        return screen_to_pdf_point(
            local_pos, self.doc[self.current_page], self.zoom
        )

    def mousePressEvent(self, event):
        if event.button() == Qt.RightButton:
            self.pan_mode = True
            self.dragged_enough = False
            self.pan_start_pos = event.globalPosition().toPoint()
            self.setCursor(Qt.ClosedHandCursor)
            event.accept()
            return

        if event.button() == Qt.LeftButton:
            self.selecting = False
            self.mouse_press_pos = event.position().toPoint()
            self.is_link_click = False

            pdf_point = self.get_pdf_point(event)
            if pdf_point and self.check_and_follow_link(pdf_point):
                self.is_link_click = True
                self.selection_start = None
                self.selection_end = None
                event.accept()
                return

            self.rubberBand.hide()
            event.accept()

    def mouseMoveEvent(self, event):
        if self.pan_mode and self.pan_start_pos is not None:
            current_pos = event.globalPosition().toPoint()
            delta = self.pan_start_pos - current_pos
            if delta.manhattanLength() > 3:
                self.dragged_enough = True

            self.horizontalScrollBar().setValue(self.horizontalScrollBar().value() + delta.x())
            self.verticalScrollBar().setValue(self.verticalScrollBar().value() + delta.y())
            self.pan_start_pos = current_pos
            event.accept()
            return

        if event.buttons() & Qt.LeftButton:
            if getattr(self, "is_link_click", False):
                event.accept()
                return

            current_pos = event.position().toPoint()
            press_pos = getattr(self, "mouse_press_pos", current_pos)

            if not self.selecting and (current_pos - press_pos).manhattanLength() > 5:
                self.selecting = True
                self.selection_start = press_pos
                self.rubberBand.setGeometry(QRect(self.selection_start, self.selection_start).normalized())
                self.rubberBand.show()

            if self.selecting:
                self.selection_end = current_pos
                self.rubberBand.setGeometry(QRect(self.selection_start, self.selection_end).normalized())
                event.accept()
                return

        if self.doc and 0 <= self.current_page < len(self.doc):
            pdf_point = self.get_pdf_point(event)
            if pdf_point:
                page = self.doc[self.current_page]
                is_over_link = any(
                    link.get("from") and link.get("from").contains(pdf_point)
                    for link in page.get_links()
                )
                self.setCursor(Qt.PointingHandCursor if is_over_link else Qt.ArrowCursor)
            else:
                self.setCursor(Qt.ArrowCursor)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.RightButton and self.pan_mode:
            self.pan_mode = False
            self.pan_start_pos = None
            self.setCursor(Qt.ArrowCursor)
            if not self.dragged_enough:
                self.show_context_menu(event.position().toPoint())
            event.accept()
            return

        if event.button() == Qt.LeftButton:
            if getattr(self, "is_link_click", False):
                self.is_link_click = False
                self.selecting = False
                event.accept()
                return

            if self.selecting:
                self.selecting = False
                self.selection_page = self.current_page
                self.selection_end = event.position().toPoint()
                self.rubberBand.hide()
                self.render_page()
            else:
                self.selection_start = None
                self.selection_end = None
                self.rubberBand.hide()
                self.render_page()

            event.accept()

    def check_and_follow_link(self, point):
        if not self.doc or not (0 <= self.current_page < len(self.doc)): return False
        page = self.doc[self.current_page]
        links = page.get_links()

        for link in links:
            rect = link.get("from")
            if not rect or not rect.contains(point): continue

            uri = link.get("uri")
            if uri:
                QDesktopServices.openUrl(QUrl(uri))
                return True

            name_action = link.get("name")
            if name_action:
                if name_action == "NextPage": self.next_page(); return True
                elif name_action == "PrevPage": self.previous_page(); return True
                elif name_action == "FirstPage": self.goto_page(0); return True
                elif name_action == "LastPage": self.goto_page(len(self.doc) - 1); return True

            target_page = link.get("page")
            to_point = link.get("to")

            if target_page is None or target_page < 0:
                named_dest = link.get("nameddest") or link.get("dest")
                if named_dest:
                    try:
                        names_dict = self.doc.resolve_names()
                        if named_dest in names_dict:
                            target_info = names_dict[named_dest]
                            target_page = target_info[0]
                    except Exception: pass

                if target_page is None and named_dest:
                    try:
                        resolved = self.doc.resolve_link(link)
                        if isinstance(resolved, (tuple, list)) and len(resolved) >= 1:
                            target_page = resolved[0]
                            if len(resolved) > 1 and resolved[1] is not None:
                                to_point = resolved[1]
                    except Exception: pass

            if target_page is not None and 0 <= target_page < len(self.doc):
                self.linkClicked.emit()
                if to_point:
                    try:
                        px = getattr(to_point, "x", to_point[0] if isinstance(to_point, (list, tuple)) else 0)
                        py = getattr(to_point, "y", to_point[1] if isinstance(to_point, (list, tuple)) else 0)
                        self.goto_position(target_page, px, py)

                    except Exception as e: print(f"[Kaydirma Hatasi]: {e}")
                elif target_page != self.current_page:
                    self.goto_page_top(target_page)
                return True
        return False

    def zoom_in(self):
        self.zoom = min(MAX_ZOOM, self.zoom + 0.2)
        self.render_page()

    def zoom_out(self):
        self.zoom = max(0.4, self.zoom - 0.2)
        self.render_page()

    def zoom_100(self):
        """PDF'i orijinal (%100) boyutuna getirir."""
        self.zoom = 1.0
        self.render_page()
        self.zoom_changed.emit()

    def fit_to_width(self):
        """PDF'i ekrandaki kaydırma alanının genişliğine tam sığdırır."""
        if not self.doc: return
        page = self.doc[self.current_page]

        # viewport() QScrollArea'nın görünür alanını verir.
        # -20 piksel ise dikey kaydırma çubuğu ve kenarlık payıdır, yatay kaydırma çubuğunun çıkmasını engeller.
        view_width = self.viewport().width() - 20
        pdf_width = page.rect.width

        if pdf_width > 0:
            self.zoom = max(0.4, min(MAX_ZOOM, view_width / pdf_width))
            self.render_page()
            self.zoom_changed.emit()
    # ---------------------------------

    def get_selected_text(self):
        if self.selection_start is None or self.selection_end is None or not self.doc: return ""
        page = self.doc[self.current_page]
        start = screen_to_pdf_point(self.selection_start, page, self.zoom)
        end = screen_to_pdf_point(self.selection_end, page, self.zoom)
        rect = pymupdf.Rect(start, end).normalize()
        return page.get_text("text", clip=rect).strip()

    def copy_selected_text(self, force_ocr=False):
        if self.selection_start is None or self.selection_end is None or not self.doc: return

        page = self.doc[self.current_page]
        start = screen_to_pdf_point(self.selection_start, page, self.zoom)
        end = screen_to_pdf_point(self.selection_end, page, self.zoom)
        rect = pymupdf.Rect(start, end).normalize()
        plain_text = page.get_text("text", clip=rect, sort=True).strip()

        if force_ocr or len(plain_text) < 3:
            try:
                matrix = pymupdf.Matrix(3, 3)
                pix = page.get_pixmap(matrix=matrix, clip=rect)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                self.ocr_thread = OCRThread(img, ocr_lang=getattr(self, "current_ocr_lang", "tur+eng"))
                self.ocr_thread.recognized.connect(self._on_ocr_finished)
                self.ocr_thread.error.connect(self._on_ocr_error)
                self.ocr_thread.start()

                self.selection_start = None
                self.selection_end = None
                self.render_page()
                return
            except Exception as e:
                print(f"[HATA] OCR Hazirlik Hatasi: {e}")

        clean_html = ""
        text_dict = page.get_text("dict", clip=rect, sort=True)
        html_blocks = []
        for block in text_dict.get("blocks", []):
            if block.get("type") == 0:
                lines = block.get("lines", [])
                if not lines: continue

                max_x1 = max(line.get("bbox", (0,0,0,0))[2] for line in lines)
                assembled_html = ""
                for i, line in enumerate(lines):
                    line_html = ""
                    for span in line.get("spans", []):
                        text = span.get("text", "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                        color_int = span.get("color", 0)
                        hex_color = f"#{color_int:06x}" if isinstance(color_int, int) else "#000000"
                        flags = span.get("flags", 0)

                        style = f"color: {hex_color};"
                        if bool(flags & 16): style += " font-weight: bold;"
                        if bool(flags & 2): style += " font-style: italic;"

                        line_html += f'<span style="{style}">{text}</span>'

                    if not line_html: continue

                    if i == 0:
                        assembled_html = line_html
                    else:
                        prev_line = lines[i-1]
                        prev_x1 = prev_line.get("bbox", (0,0,0,0))[2]
                        if max_x1 - prev_x1 < 30: assembled_html += " " + line_html
                        else: assembled_html += f"</p><p style='margin: 0; text-align: left;'>{line_html}"

                if assembled_html:
                    html_blocks.append(f"<p style='margin: 0 0 8px 0; text-align: left;'>{assembled_html}</p>")

        clean_html = "\n".join(html_blocks)

        mime_data = QMimeData()
        safe_plain_text = plain_text.replace('\n', ' \r\n').replace('  \r\n', ' \r\n')
        mime_data.setText(safe_plain_text)
        full_html = f"<html><body><div style='text-align: left;'>{clean_html}</div></body></html>"
        mime_data.setHtml(full_html)
        QApplication.clipboard().setMimeData(mime_data)

        self.selection_start = None
        self.selection_end = None
        self.render_page()

    def _on_ocr_finished(self, plain_text):
        if not plain_text: return

        mime_data = QMimeData()
        normalized_ocr = plain_text.replace('\r\n', '\n')
        paragraphs = normalized_ocr.split('\n\n')

        html_lines = []
        plain_lines = []
        for p in paragraphs:
            lines = [l.strip() for l in p.split('\n') if l.strip()]
            if not lines: continue

            max_len = max(len(l) for l in lines)
            assembled_p = lines[0]
            assembled_plain = lines[0]

            for i in range(1, len(lines)):
                prev_len = len(lines[i-1])

                if max_len > 40 and (max_len - prev_len) < 15:
                    assembled_p += " " + lines[i]
                    assembled_plain += " " + lines[i]
                else:
                    assembled_p += f"</p><p style='margin: 0; text-align: left;'>{lines[i]}"
                    assembled_plain += "\r\n" + lines[i]

            html_lines.append(f"<p style='margin: 0 0 8px 0; text-align: left;'>{assembled_p}</p>")
            plain_lines.append(assembled_plain)

        html_formatted = "\n".join(html_lines)
        full_html = f"<html><body><div style='font-family: sans-serif;'>{html_formatted}</div></body></html>"
        safe_plain_text = "\r\n\r\n".join(plain_lines)

        mime_data.setText(safe_plain_text)
        mime_data.setHtml(full_html)
        QApplication.clipboard().setMimeData(mime_data)

    def _on_ocr_error(self, err_msg):
        print(f"[HATA] OCR Hatasi: {err_msg}")

    def create_bookmark_from_selection(self):
        text = self.get_selected_text()
        if not text: return
        title = " ".join(text.split())
        if len(title) > 60: title = title[:60] + "..."
        pos = self.get_selection_position()
        if pos is None: return
        self.bookmarkRequested.emit(self.current_page, title, pos[0], pos[1])

    def create_highlight(self):
        if not getattr(self, "current_file_path", "").lower().endswith(".pdf"): return
        if self.selection_start is None or self.selection_end is None: return

        page = self.doc[self.current_page]

        if self.selection_start == self.selection_end:
            return
        start = screen_to_pdf_point(self.selection_start, page, self.zoom)
        end = screen_to_pdf_point(self.selection_end, page, self.zoom)
        pdf_rect = pymupdf.Rect(start, end).normalize()

        self.highlightRequested.emit(self.current_page, pdf_rect.x0, pdf_rect.y0, pdf_rect.x1, pdf_rect.y1)

    def export_table_from_selection(self):
        if self.selection_start is None or self.selection_end is None: return
        page = self.doc[self.current_page]
        start = screen_to_pdf_point(self.selection_start, page, self.zoom)
        end = screen_to_pdf_point(self.selection_end, page, self.zoom)
        rect = pymupdf.Rect(start, end).normalize()
        self.exportTableRequested.emit(
            self.current_page, rect.x0, rect.y0, rect.x1, rect.y1
        )

    def get_selection_position(self):
        if self.selection_start is None or self.selection_end is None: return None
        page = self.doc[self.current_page]
        point = screen_to_pdf_point(self.selection_start, page, self.zoom)
        return point.x, point.y

    def show_context_menu(self, pos):
        menu = QMenu(self)
        copy_action = menu.addAction(t("view_copy"))
        copy_ocr_action = menu.addAction(t("view_copy_ocr"))
        bookmark_action = menu.addAction(t("view_add_bookmark"))
        highlight_action = menu.addAction(t("view_add_highlight"))
        delete_highlight_action = menu.addAction(t("view_del_highlight"))
        export_table_action = menu.addAction(t("view_export_excel"))

        # DÜZELTME: Menü metni dil dosyasından (lang.py) çekiliyor
        add_note_action = menu.addAction(t("menu_add_note"))

        is_pdf = getattr(self, "current_file_path", "").lower().endswith(".pdf")
        highlight_action.setEnabled(is_pdf)
        delete_highlight_action.setEnabled(is_pdf)
        add_note_action.setEnabled(is_pdf)

        selected = menu.exec(self.image_label.mapToGlobal(pos))

        if selected == copy_action:
            self.copy_selected_text(force_ocr=False)
        elif selected == copy_ocr_action:
            self.copy_selected_text(force_ocr=True)
        elif selected == bookmark_action:
            self.create_bookmark_from_selection()
        elif selected == highlight_action:
            self.create_highlight()
        elif selected == export_table_action:
            self.export_table_from_selection()
        elif selected in (add_note_action, delete_highlight_action):
            page = self.doc[self.current_page]
            pdf_p = screen_to_pdf_point(pos, page, self.zoom)

            if selected == add_note_action:
                self.addNoteRequested.emit(self.current_page, pdf_p.x, pdf_p.y)
            elif selected == delete_highlight_action:
                self.deleteHighlightRequested.emit(self.current_page, pdf_p.x, pdf_p.y)

    def wheelEvent(self, event):
        if event.modifiers() & Qt.ControlModifier:
            if event.angleDelta().y() > 0: self.zoom_in()
            else: self.zoom_out()
            self.zoom_changed.emit()
            event.accept()
            return

        v_scrollbar = self.verticalScrollBar()
        delta = event.angleDelta().y()

        if delta < 0:
            if v_scrollbar.value() < v_scrollbar.maximum():
                super().wheelEvent(event)
            else:
                if self.doc and self.current_page < len(self.doc) - 1:
                    self.next_page()
                    v_scrollbar.setValue(0)
                event.accept()
        elif delta > 0:
            if v_scrollbar.value() > 0:
                super().wheelEvent(event)
            else:
                if self.current_page > 0:
                    self.previous_page()
                    v_scrollbar.setValue(v_scrollbar.maximum())
                event.accept()

