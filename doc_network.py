#!/usr/bin/env python3
import datetime
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import pymupdf
from PySide6.QtCore import Qt, QStandardPaths, QThread, Signal
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QApplication, QDialog, QFileDialog, QHBoxLayout, QInputDialog, QLabel,
    QLineEdit, QListWidget, QListWidgetItem, QMainWindow, QMenu,
    QMessageBox, QPushButton, QSplitter, QStackedWidget, QTextBrowser,
    QTreeWidgetItem, QVBoxLayout, QWidget, QComboBox, QTabWidget, QProgressDialog,
    QScrollArea, QGridLayout
)
from PySide6.QtPrintSupport import QPrinter, QPrintDialog
from PySide6.QtWebEngineCore import QWebEnginePage

from database import Database, get_user_data_dir, migrate_legacy_database
from widgets.bookmark_tree import BookmarkTree
from widgets.doc_viewer import DOCViewer
from lang import t, set_lang
from widgets.recent_file_card import RecentFileCard
from widgets.doc_workers import GlobalLibrarySearchWorker, PDFPrintWorker, PDFSearchWorker
from widgets.office_cache import get_office_cache_path
from widgets.export_manager import ExportManager
from widgets.md_viewer import MarkdownViewer
from widgets.data_backup import create_data_backup, restore_data_backup
from widgets.settings_dialog import SettingsDialog

class OfficeConvertWorker(QThread):
    converted = Signal(str, str, int, float, float)
    error = Signal(str)

    def __init__(self, input_path, page, x, y):
        super().__init__()
        self.input_path = input_path
        self.page = page
        self.x = x
        self.y = y

    def run(self):
        try:
            input_p = Path(self.input_path).resolve()
            cache_dir = Path(tempfile.gettempdir()) / "Doc_Network_Cache"

            cache_dir.mkdir(parents=True, exist_ok=True)

            cached_pdf = get_office_cache_path(input_p, cache_dir)

            if cached_pdf.exists() and cached_pdf.stat().st_size > 0:
                self.converted.emit(str(cached_pdf), self.input_path, self.page, self.x, self.y)
                return

            with tempfile.TemporaryDirectory(dir=cache_dir) as temp_dir:
                temp_working_dir = Path(temp_dir)
                temp_input_file = temp_working_dir / input_p.name
                shutil.copy2(str(input_p), str(temp_input_file))

                expected_temp_pdf = temp_working_dir / f"{input_p.stem}.pdf"
                converted = False

                if sys.platform == "win32" and input_p.suffix.lower() == ".docx":
                    try:
                        import win32com.client
                        import pythoncom
                        pythoncom.CoInitialize()
                        word = win32com.client.DispatchEx("Word.Application")
                        word.Visible = False
                        word_doc = word.Documents.Open(str(temp_input_file))
                        word_doc.SaveAs(str(expected_temp_pdf), FileFormat=17)
                        word_doc.Close()
                        word.Quit()
                        pythoncom.CoUninitialize()

                        if expected_temp_pdf.exists() and expected_temp_pdf.stat().st_size > 0:
                            converted = True
                    except Exception:
                        pass

                if not converted:
                    lo_profile = temp_working_dir / "lo_profile"
                    env_arg = f"-env:UserInstallation=file://{lo_profile.as_posix()}"

                    pdf_export_filter = "writer_pdf_Export"
                    filter_data = ""

                    if input_p.suffix.lower() in [".xlsx", ".xls", ".ods"]:
                        pdf_export_filter = "calc_pdf_Export"
                        filter_data = '{"FitWidth":{v:"1",type:"long"},"IsFitToPage":{v:"true",type:"boolean"}}'

                    cmd = [
                        "soffice", env_arg, "--headless",
                        "--convert-to", f"pdf:{pdf_export_filter}:{filter_data}",
                        "--outdir", str(temp_working_dir),
                        str(temp_input_file)
                    ]
                    subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=25)

                if expected_temp_pdf.exists() and expected_temp_pdf.stat().st_size > 0:
                    shutil.move(str(expected_temp_pdf), str(cached_pdf))
                else:
                    self.error.emit("Dönüştürülen PDF boş veya oluşturulamadı.")
                    return

            self.converted.emit(str(cached_pdf), self.input_path, self.page, self.x, self.y)

        except Exception as e:
            self.error.emit(str(e))

class NoteBrowser(QTextBrowser):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.main_window = None

    def mouseDoubleClickEvent(self, event):
        if self.main_window:
            self.main_window.enable_note_editing()
        super().mouseDoubleClickEvent(event)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        if "APPIMAGE" in os.environ:
            self.proje_ana_dizini = Path(os.environ["APPIMAGE"]).parent
        else:
            self.proje_ana_dizini = Path(os.path.realpath(__file__)).parent

        self.user_data_dir = get_user_data_dir()
        self.user_data_dir.mkdir(parents=True, exist_ok=True)
        legacy_data_dir = self.proje_ana_dizini / "data"
        migrate_legacy_database(
            legacy_data_dir / "bookmarks.db",
            self.user_data_dir / "bookmarks.db",
        )
        legacy_settings = legacy_data_dir / "settings.json"
        self.settings_file = self.user_data_dir / "settings.json"
        if not self.settings_file.exists() and legacy_settings.is_file():
            shutil.copy2(legacy_settings, self.settings_file)

        self.db = Database(self.user_data_dir / "bookmarks.db")
        self.export_manager = ExportManager(self)
        self.current_theme = "light"
        documents_dir = QStandardPaths.writableLocation(
            QStandardPaths.DocumentsLocation
        )
        documents_path = Path(documents_dir) if documents_dir else Path.home()
        self.default_library_path = (
            documents_path if documents_path.is_dir() else Path.home()
        )
        self.library_path = self.default_library_path
        self.last_opened_dir = None

        self.resource_dir = Path(
            getattr(sys, "_MEIPASS", Path(__file__).resolve().parent)
        )
        icon_path = self.resource_dir / "icon.png"
        if not icon_path.exists():
            icon_path = self.resource_dir / "icon.svg"
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))

        self.tree = BookmarkTree()
        self.search_worker = None
        self.pdf_search_workers = []
        self.global_search_workers = []
        self.search_request_id = 0

        self.page_label = QLabel()
        self.zoom_label = QLabel()
        self.btn_open = QPushButton()
        self.btn_group = QPushButton()
        self.btn_prev = QPushButton()
        self.btn_next = QPushButton()
        self.btn_zoom_in = QPushButton()
        self.btn_zoom_out = QPushButton()

        self.btn_toggle = QPushButton()
        self.btn_zoom_100 = QPushButton("%100")
        self.btn_zoom_fit = QPushButton("↔ Sığdır")

        self.btn_print = QPushButton()
        self.btn_export = QPushButton()
        self.note_label = QLabel()
        self.btn_save_note = QPushButton()
        self.btn_search = QPushButton()
        self.dash_title = QLabel()
        self.dash_recent_lbl = QLabel()
        self.btn_dashboard = QPushButton()

        self.btn_nav_back = QPushButton("⬅️")
        self.btn_nav_forward = QPushButton("➡️")
        self.nav_back_stack = []
        self.nav_forward_stack = []
        self.is_navigating_history = False

        self.btn_theme_toggle = QPushButton()
        self.btn_settings = QPushButton()


        self.selected_group_id = None
        self.active_bookmark_id = None
        self.pdf_history = []
        self.current_pdf = None

        self.setup_ui()
        self.load_settings()
        self.metinleri_guncelle()
        self.load_groups()

        self.tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self.show_tree_menu)
        self.tree.leftClickedItem.connect(self.on_tree_item_clicked)


    def get_current_viewer(self):
        if self.tab_widget.count() > 0:
            return self.tab_widget.currentWidget()
        return None

    def metinleri_guncelle(self):
        base_title = t("app_title")
        if self.current_pdf:
            self.setWindowTitle(f"{base_title} - {Path(self.current_pdf).name}")
        else:
            self.setWindowTitle(base_title)

        self.update_page_label()
        self.update_zoom_label()

        self.btn_dashboard.setText(t("btn_dashboard"))
        self.update_dashboard_button()
        self.page_edit.setPlaceholderText(t("ph_page"))
        self.btn_open.setText(t("btn_open"))
        self.btn_group.setText(t("btn_group"))
        self.btn_prev.setText(t("btn_prev"))
        self.btn_next.setText(t("btn_next"))
        self.btn_zoom_in.setText(t("btn_zoom_in"))
        self.btn_zoom_out.setText(t("btn_zoom_out"))
        self.btn_toggle.setText(t("btn_toggle"))
        self.btn_print.setText(t("btn_print"))
        self.btn_export.setText(t("btn_export"))
        self.btn_settings.setText(t("btn_settings"))

        self.search_box.setPlaceholderText(t("ph_search_tree"))
        self.note_label.setText(t("lbl_note"))
        self.note_edit.setPlaceholderText(t("ph_note"))
        self.btn_save_note.setText(t("btn_save_note"))

        self.pdf_search.setPlaceholderText(t("ph_search_pdf"))
        self.btn_search.setText(t("btn_search_find"))
        self.search_mode.setItemText(0, t("search_open_document"))
        self.search_mode.setItemText(1, t("search_dashboard_tabs"))
        self.search_mode.setItemText(2, t("search_library_folder"))

        self.dash_title.setText(t("dash_title"))
        self.dash_recent_lbl.setText(t("dash_recent_title"))
        self.tree.setHeaderLabel(t("tree_header"))

        self.update_dashboard_ui(self.pdf_history)
        self.load_groups()

        self.btn_toggle.setText(t("btn_toggle"))

        self.btn_zoom_100.setText("%100")
        if self.combo_lang.currentIndex() == 1: # İngilizce ise
            self.btn_zoom_fit.setText("↔ Fit Width")
        else:
            self.btn_zoom_fit.setText("↔ Sığdır")

    def lang_degistir(self, index):
        set_lang("tr" if index == 0 else "en")
        self.metinleri_guncelle()
        self.save_settings()

    def get_dark_stylesheet(self):
        return """
        QMainWindow, QWidget { background-color: #151a21; color: #d1d6dc; }
        QSplitter::handle { background: #232b35; }
        QLineEdit, QTextBrowser, QTreeView, QListView, QTableView, QScrollArea {
            background-color: #1a2027;
            border: 1px solid #2d3640;
            color: #d1d6dc;
            border-radius: 4px;
            alternate-background-color: #1e2630;
        }
        QPushButton {
            background-color: #1e2630; border: 1px solid #2d3640; border-radius: 4px; padding: 4px 8px; color: #d1d6dc;
        }
        QPushButton:hover { background-color: #26303c; border-color: #38434f; }
        QPushButton:pressed { background-color: #151a21; }
        QPushButton:disabled { color: #5a6673; background-color: #151a21; border-color: #1e2630; }
        QComboBox { background-color: #1e2630; border: 1px solid #2d3640; color: #d1d6dc; border-radius: 4px; padding: 2px 8px; }
        QComboBox::drop-down { border: none; }
        QTabWidget::pane { border: 1px solid #2d3640; background: #151a21; }
        QTabBar::tab { background-color: #1e2630; color: #8c97a1; padding: 6px 12px; border: 1px solid #2d3640; border-bottom: none; border-top-left-radius: 4px; border-top-right-radius: 4px; margin-right: 2px; }
        QTabBar::tab:selected { background-color: #151a21; color: #d1d6dc; font-weight: bold; }
        QTreeView::item:selected, QListView::item:selected, QTableView::item:selected { background-color: #137886; color: white; }
        QTreeView::item:hover, QListView::item:hover, QTableView::item:hover { background-color: #1e2630; }
        QHeaderView::section { background-color: #1e2630; color: #d1d6dc; border: 1px solid #2d3640; padding: 4px; }
        QMenu { background-color: #1a2027; color: #d1d6dc; border: 1px solid #2d3640; }
        QMenu::item:selected { background-color: #137886; }
        """

    def apply_theme(self):
        if self.current_theme == "dark":
            self.btn_theme_toggle.setText("☀️ Light")
            QApplication.instance().setStyleSheet(self.get_dark_stylesheet())
            self.dash_title.setStyleSheet("font-size: 20px; font-weight: bold; color: #d1d6dc;")
            self.stats_label.setStyleSheet("font-size: 13px; color: #8c97a1; font-style: italic;")
            self.dash_recent_lbl.setStyleSheet("color: #d1d6dc;")
            self.note_label.setStyleSheet("font-weight: bold; margin-top: 5px; color: #d1d6dc;")
        else:
            self.btn_theme_toggle.setText("🌙 Dark")
            QApplication.instance().setStyleSheet("")
            self.dash_title.setStyleSheet("font-size: 20px; font-weight: bold; color: #2c3e50;")
            self.stats_label.setStyleSheet("font-size: 13px; color: #7f8c8d; font-style: italic;")
            self.dash_recent_lbl.setStyleSheet("color: #000000;")
            self.note_label.setStyleSheet("font-weight: bold; margin-top: 5px; color: #000000;")

        self.update_dashboard_ui(self.pdf_history)

        for i in range(self.tab_widget.count()):
            w = self.tab_widget.widget(i)
            if isinstance(w, MarkdownViewer):
                w.theme = self.current_theme
                w.open_markdown(w.current_file_path)

    def toggle_theme(self):
        self.current_theme = "dark" if self.current_theme == "light" else "light"
        self.apply_theme()
        self.save_settings()

    def setup_ui(self):
        central = QWidget()
        root = QVBoxLayout()
        toolbar = QHBoxLayout()

        self.page_edit = QLineEdit()
        self.page_edit.setFixedWidth(80)
        self.page_edit.returnPressed.connect(self.goto_page)

        toolbar.setContentsMargins(0, 0, 0, 0)
        toolbar.setSpacing(4)

        self.btn_dashboard.clicked.connect(self.toggle_dashboard)

        self.btn_nav_back.setText("◀")
        self.btn_nav_forward.setText("▶")
        self.btn_nav_back.clicked.connect(self.navigate_back)
        self.btn_nav_forward.clicked.connect(self.navigate_forward)
        self.btn_nav_back.setEnabled(False)
        self.btn_nav_forward.setEnabled(False)
        self.btn_nav_back.setMinimumWidth(36)
        self.btn_nav_forward.setMinimumWidth(36)

        self.btn_open.clicked.connect(self.open_pdf)
        self.btn_prev.clicked.connect(self.previous_page)
        self.btn_next.clicked.connect(self.next_page)
        self.btn_zoom_in.clicked.connect(self.zoom_in)
        self.btn_zoom_out.clicked.connect(self.zoom_out)

        self.btn_zoom_100.clicked.connect(self.zoom_100)
        self.btn_zoom_fit.clicked.connect(self.zoom_fit)

        self.btn_group.clicked.connect(self.add_group)
        self.btn_toggle.clicked.connect(self.toggle_bookmarks)
        self.btn_print.clicked.connect(self.print_pdf)
        self.btn_export.clicked.connect(
            lambda: self.export_manager.export_bookmarks_to_markdown(self.proje_ana_dizini)
        )

        self.combo_lang = QComboBox()
        self.combo_lang.addItems(["🇹🇷 TR", "🇬🇧 EN"])
        self.combo_lang.setMinimumWidth(85)
        self.combo_lang.currentIndexChanged.connect(self.lang_degistir)

        self.btn_theme_toggle.clicked.connect(self.toggle_theme)
        self.btn_settings.clicked.connect(self.open_settings_dialog)

        for btn in (self.btn_dashboard, self.btn_nav_back, self.btn_nav_forward, self.btn_print, self.btn_open, self.btn_prev, self.btn_next,
                    self.btn_zoom_in, self.btn_zoom_out, self.btn_zoom_100, self.btn_zoom_fit, self.btn_group,
                    self.btn_toggle, self.btn_export, self.page_label,
                    self.zoom_label, self.combo_lang, self.btn_theme_toggle,
                    self.btn_settings):
            btn.setMaximumHeight(28)

        # --- BUTON YERLEŞİM SIRASI ---
        toolbar.addWidget(self.btn_toggle)
        toolbar.addWidget(self.btn_dashboard)

        toolbar.addWidget(self.btn_nav_back)
        toolbar.addWidget(self.btn_nav_forward)

        toolbar.addWidget(self.btn_open)
        toolbar.addWidget(self.btn_print)
        toolbar.addWidget(self.btn_export)
        toolbar.addWidget(self.btn_group)
        toolbar.addWidget(self.btn_prev)
        toolbar.addWidget(self.btn_next)
        toolbar.addWidget(self.page_edit)

        toolbar.addWidget(self.btn_zoom_in)
        toolbar.addWidget(self.btn_zoom_out)
        toolbar.addWidget(self.btn_zoom_100)
        toolbar.addWidget(self.btn_zoom_fit)

        toolbar.addStretch()
        toolbar.addWidget(self.zoom_label)
        toolbar.addWidget(self.page_label)
        toolbar.addWidget(self.combo_lang)
        toolbar.addWidget(self.btn_theme_toggle)
        toolbar.addWidget(self.btn_settings)
        # ------------------------------------------

        self.search_box = QLineEdit()
        self.search_box.textChanged.connect(self.filter_tree)

        self.left_panel = QWidget()
        self.left_panel.setMinimumWidth(200)
        left_layout = QVBoxLayout(self.left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(4)

        self.note_edit = NoteBrowser()
        self.note_edit.main_window = self
        self.note_edit.setReadOnly(True)
        self.note_edit.setOpenLinks(False)
        self.note_edit.setMaximumHeight(150)
        self.note_edit.setEnabled(False)
        self.note_edit.anchorClicked.connect(self.handle_knowledge_link)

        self.btn_save_note.setMaximumHeight(28)
        self.btn_save_note.setEnabled(False)
        self.btn_save_note.clicked.connect(self.save_current_bookmark_note)

        left_layout.addWidget(self.search_box)
        left_layout.addWidget(self.tree, 1)
        left_layout.addWidget(self.note_label)
        left_layout.addWidget(self.note_edit)
        left_layout.addWidget(self.btn_save_note)

        self.search_results = QListWidget()
        self.search_results.itemDoubleClicked.connect(self.goto_search_result)

        self.search_mode = QComboBox()
        self.search_mode.addItems([
            t("search_open_document"),
            t("search_dashboard_tabs"),
            t("search_library_folder"),
        ])

        self.pdf_search = QLineEdit()
        self.pdf_search.returnPressed.connect(self.search_in_pdf)
        self.btn_search.clicked.connect(self.search_in_pdf)

        self.search_results = QListWidget()
        self.search_results.itemDoubleClicked.connect(self.goto_search_result)

        search_panel = QWidget()
        search_layout = QVBoxLayout(search_panel)
        search_layout.setContentsMargins(0, 0, 0, 0)
        search_layout.addWidget(self.search_mode)
        search_layout.addWidget(self.pdf_search)
        search_layout.addWidget(self.btn_search)
        search_layout.addWidget(self.search_results)


        self.splitter = QSplitter()
        self.splitter.addWidget(self.left_panel)

        self.right_stack = QStackedWidget()

        self.dashboard_widget = QWidget()
        dash_layout = QVBoxLayout(self.dashboard_widget)
        dash_layout.setContentsMargins(20, 20, 20, 20)
        dash_layout.setSpacing(15)

        dash_layout.addWidget(self.dash_title)

        self.stats_label = QLabel()
        dash_layout.addWidget(self.stats_label)

        dash_layout.addWidget(self.dash_recent_lbl)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setStyleSheet("background-color: transparent; border: none;")

        self.grid_widget = QWidget()
        self.grid_widget.setStyleSheet("background-color: transparent;")
        self.grid_layout = QGridLayout(self.grid_widget)
        self.grid_layout.setSpacing(15)
        self.grid_layout.setAlignment(Qt.AlignLeft | Qt.AlignTop)

        self.scroll_area.setWidget(self.grid_widget)
        dash_layout.addWidget(self.scroll_area, 1)

        self.tab_widget = QTabWidget()
        self.tab_widget.setTabsClosable(True)
        self.tab_widget.tabCloseRequested.connect(self.close_tab)
        self.tab_widget.currentChanged.connect(self.on_tab_changed)

        self.right_stack.addWidget(self.dashboard_widget)
        self.right_stack.addWidget(self.tab_widget)
        self.right_stack.currentChanged.connect(self.update_dashboard_button)

        self.splitter.addWidget(self.right_stack)
        self.splitter.addWidget(search_panel)

        self.splitter.setSizes([250, 900, 300])
        self.splitter.setStretchFactor(0, 0)
        self.splitter.setStretchFactor(1, 1)

        root.addLayout(toolbar)
        root.addWidget(self.splitter, 1)

        central.setLayout(root)
        self.setCentralWidget(central)
        self.resize(1400, 900)

    def on_tab_changed(self, index):
        viewer = self.get_current_viewer()
        if viewer:
            self.current_pdf = getattr(viewer, "original_file_path", viewer.current_file_path)
            self.setWindowTitle(f"{t('app_title')} - {Path(self.current_pdf).name}")
            self.update_page_label()
            self.update_zoom_label()
            self.load_pdf_bookmarks()
            self.btn_zoom_in.setEnabled(True)
            self.btn_zoom_out.setEnabled(True)
        else:
            self.current_pdf = None
            self.setWindowTitle(t("app_title"))
            self.page_label.setText(t("lbl_page_empty"))
            self.page_edit.clear()
            self.btn_zoom_in.setEnabled(False)
            self.btn_zoom_out.setEnabled(False)

            # AÇIK BELGE KALMADIĞINDA BELGE YER İMLERİNİ TEMİZLE
            self.load_pdf_bookmarks()

        self.update_dashboard_button()

    def on_tree_item_clicked(self, item):
        group_id = item.data(0, Qt.UserRole)
        if group_id is not None:
            self.selected_group_id = group_id
        self.open_bookmark(item)

    def toggle_dashboard(self):
        if self.right_stack.currentIndex() == 1:
            self.right_stack.setCurrentIndex(0)
            self.setWindowTitle(t("app_title"))
        else:
            if self.tab_widget.count() > 0:
                self.right_stack.setCurrentIndex(1)
                if self.current_pdf:
                    self.setWindowTitle(f"{t('app_title')} - {Path(self.current_pdf).name}")

    def update_dashboard_button(self, index=None):
        if self.right_stack.currentIndex() == 0:
            if self.tab_widget.count() > 0:
                self.btn_dashboard.setText(t("btn_back_to_doc"))
                self.btn_dashboard.setEnabled(True)
            else:
                self.btn_dashboard.setText(t("btn_dashboard"))
                self.btn_dashboard.setEnabled(False)
        else:
            self.btn_dashboard.setText(t("btn_dashboard"))
            self.btn_dashboard.setEnabled(True)

    def record_navigation_point(self):
        if self.is_navigating_history:
            return
        viewer = self.get_current_viewer()
        if not viewer or not getattr(viewer, "doc", None):
            return
        current_path = getattr(viewer, "original_file_path", viewer.current_file_path)
        current_page = viewer.current_page

        if isinstance(viewer, DOCViewer):
            x, y = viewer.get_pdf_position_from_scroll()
        else:
            x = 0.0
            y = getattr(viewer, "current_scroll_y", 0.0)

        point = (current_path, current_page, x, y)
        # Aynı dosyada bile scrollY (y) değişmişse yeni navigasyon noktası olarak kaydet!
        if self.nav_back_stack and self.nav_back_stack[-1][:2] == (current_path, current_page):
            if abs(self.nav_back_stack[-1][3] - y) < 20:
                return

        self.nav_back_stack.append(point)
        if len(self.nav_back_stack) > 50:
            self.nav_back_stack.pop(0)
        self.nav_forward_stack.clear()
        self.update_nav_buttons()

    def _get_current_nav_point(self):
        viewer = self.get_current_viewer()
        if not viewer or not getattr(viewer, "doc", None):
            return None
        current_path = getattr(viewer, "original_file_path", viewer.current_file_path)

        if isinstance(viewer, DOCViewer):
            x, y = viewer.get_pdf_position_from_scroll()
        else:
            x = 0.0
            y = getattr(viewer, "current_scroll_y", 0.0)

        return (current_path, viewer.current_page, x, y)


    def update_nav_buttons(self):
        self.btn_nav_back.setEnabled(len(self.nav_back_stack) > 0)
        self.btn_nav_forward.setEnabled(len(self.nav_forward_stack) > 0)

    def navigate_back(self):
        if not self.nav_back_stack:
            return
        current_pt = self._get_current_nav_point()
        if current_pt:
            self.nav_forward_stack.append(current_pt)

        target_path, target_page, x, y = self.nav_back_stack.pop()
        self.is_navigating_history = True
        try:
            self.open_in_tab(target_path, target_page, x, y)
        finally:
            self.is_navigating_history = False
            self.update_nav_buttons()

    def navigate_forward(self):
        if not self.nav_forward_stack:
            return
        current_pt = self._get_current_nav_point()
        if current_pt:
            self.nav_back_stack.append(current_pt)

        target_path, target_page, x, y = self.nav_forward_stack.pop()
        self.is_navigating_history = True
        try:
            self.open_in_tab(target_path, target_page, x, y)
        finally:
            self.is_navigating_history = False
            self.update_nav_buttons()


    def add_to_history(self, pdf_path, page):
        history = getattr(self, "pdf_history", [])
        history = [item for item in history if (item.get("path") if isinstance(item, dict) else item) != pdf_path]
        history.insert(0, {
            "path": pdf_path,
            "page": page,
            "date": datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        })
        self.pdf_history = history[:5]
        self.update_dashboard_ui(self.pdf_history)

    def close_tab(self, index):
        viewer = self.tab_widget.widget(index)
        if viewer:
            for worker in (
                getattr(viewer, "_render_worker", None),
                getattr(viewer, "ocr_thread", None),
            ):
                if worker is not None and worker.isRunning():
                    worker.requestInterruption()
                    worker.wait()
            if hasattr(viewer, "doc") and viewer.doc:
                viewer.doc.close()
            viewer.deleteLater()
        self.tab_widget.removeTab(index)
        if self.tab_widget.count() == 0:
            self.right_stack.setCurrentIndex(0)

    def open_in_tab(self, file_path, page=0, x=0, y=None):
        opened_dir = Path(file_path).expanduser().resolve().parent
        if opened_dir.is_dir():
            self.last_opened_dir = str(opened_dir)

        for i in range(self.tab_widget.count()):
            viewer = self.tab_widget.widget(i)
            if getattr(viewer, "original_file_path", viewer.current_file_path) == file_path:
                self.tab_widget.setCurrentIndex(i)
                if y is not None or page > 0:
                    if x == 0 and not y and isinstance(viewer, DOCViewer):
                        viewer.goto_page_top(page)
                    else:
                        viewer.goto_position(page, x, y or 0)
                self.right_stack.setCurrentIndex(1)
                self.add_to_history(file_path, page)
                self.save_settings()
                return i

        suffix = Path(file_path).suffix.lower()

        # Markdown dosyalarını doğrudan canlı web görünümünde aç (GIF ve tam genişlik desteği)
        if suffix in [".md", ".markdown"]:
            md_viewer = MarkdownViewer(theme=self.current_theme)
            md_viewer.zoom_changed.connect(self.update_zoom_label)
            md_viewer.bookmarkRequested.connect(self.add_bookmark_from_selection)
            md_viewer.open_markdown(file_path)
            md_viewer.linkClicked.connect(self.record_navigation_point)

            if y is not None and (page > 0 or y > 0): md_viewer.goto_position(page, x, y)

            tab_title = Path(file_path).name
            index = self.tab_widget.addTab(md_viewer, tab_title)
            self.tab_widget.setCurrentIndex(index)
            self.right_stack.setCurrentIndex(1)
            self.add_to_history(file_path, page)
            self.save_settings()
            self.load_pdf_bookmarks()
            return index

        is_office_doc = suffix in [".docx", ".odt", ".xlsx", ".ods", ".pptx", ".ppt", ".odp", ".odg", ".odf", ".rtf", ".csv"]

        if is_office_doc:
            self.convert_progress = QProgressDialog(t("lbl_loading"), None, 0, 0, self)
            self.convert_progress.setWindowTitle(t("lbl_loading"))
            self.convert_progress.setWindowModality(Qt.WindowModal)
            self.convert_progress.show()

            self.office_worker = OfficeConvertWorker(file_path, page, x, y)
            self.office_worker.converted.connect(self._on_office_converted)
            self.office_worker.error.connect(self._on_office_error)
            self.office_worker.start()
            return -1

        self._load_pdf_into_tab(file_path, file_path, False, page, x, y)
        return self.tab_widget.currentIndex()
    def _on_office_converted(self, pdf_path, original_path, page, x, y):
        if hasattr(self, "convert_progress"):
            self.convert_progress.close()
        self._load_pdf_into_tab(pdf_path, original_path, True, page, x, y)

    def _on_office_error(self, err_msg):
        if hasattr(self, "convert_progress"):
            self.convert_progress.close()
        QMessageBox.critical(self, t("err_convert_title"), t("err_convert_msg") + f"\n\nSistem Detayı: {err_msg}")

    def _load_pdf_into_tab(self, active_open_path, original_path, is_office_doc, page, x, y):
        new_viewer = DOCViewer()
        new_viewer.original_file_path = original_path
        new_viewer.is_office_doc = is_office_doc

        new_viewer.zoom_changed.connect(self.update_zoom_label)
        new_viewer.bookmarkRequested.connect(self.add_bookmark_from_selection)
        new_viewer.highlightRequested.connect(self.create_highlight)
        new_viewer.deleteHighlightRequested.connect(self.delete_highlight)
        new_viewer.exportTableRequested.connect(
            lambda p, x0, y0, x1, y1: self.export_manager.export_table_to_excel(new_viewer, p, x0, y0, x1, y1)
        )
        new_viewer.addNoteRequested.connect(self.add_page_note)
        new_viewer.page_changed.connect(self.update_page_label)
        new_viewer.linkClicked.connect(self.record_navigation_point)
        new_viewer.renderError.connect(
            lambda message: QMessageBox.warning(self, t("msg_error"), message)
        )

        new_viewer.open_pdf(active_open_path)
        if x == 0 and not y:
            new_viewer.goto_page_top(page)
        else:
            new_viewer.goto_position(page, x, y or 0)

        tab_title = Path(original_path).name
        index = self.tab_widget.addTab(new_viewer, tab_title)
        self.tab_widget.setCurrentIndex(index)
        self.right_stack.setCurrentIndex(1)
        self.add_to_history(original_path, page)
        self.save_settings()
        return index

    def open_pdf(self):
        candidate_dirs = (
            self.last_opened_dir,
            self.library_path,
            QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation),
            Path.home(),
        )
        default_dir = next(
            str(Path(directory).expanduser())
            for directory in candidate_dirs
            if directory and Path(directory).expanduser().is_dir()
        )
        file_filter = (
            "Tüm Desteklenen Belgeler (*.pdf *.xlsx *.ods *.docx *.odt *.pptx *.odp *.odg *.odf *.epub *.mobi *.fb2 *.cbz *.txt *.rtf *.csv);;"
            "PDF ve E-Kitaplar (*.pdf *.epub *.mobi *.fb2 *.cbz);;"
            "Desteklenen Ofis ve PDF Belgeleri (*.pdf *.ods *.odt *.odp *.odg *.odf *.xlsx *.docx *.pptx);;"
            "Markdown Dosyaları (*.md *.markdown);;"
            "Tüm Dosyalar (*.*)"
        )
        filename, _ = QFileDialog.getOpenFileName(self, t("btn_open"), default_dir, file_filter)
        if filename:
            self.open_in_tab(filename)

    def open_settings_dialog(self):
        dialog = SettingsDialog(self.library_path, self)
        dialog.backup_export_requested.connect(self.export_data_backup)
        dialog.backup_import_requested.connect(
            lambda: self.import_data_backup(dialog)
        )
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.library_path = dialog.selected_library_path
            self.save_settings()

    def export_data_backup(self):
        self.save_settings()
        default_name = (
            f"Doc_Network-backup-{datetime.datetime.now():%Y%m%d-%H%M%S}.zip"
        )
        filename, _ = QFileDialog.getSaveFileName(
            self,
            t("settings_export_title"),
            str(self.user_data_dir / default_name),
            t("settings_zip_filter"),
        )
        if not filename:
            return

        archive_path = Path(filename)
        if archive_path.suffix.lower() != ".zip":
            archive_path = archive_path.with_suffix(".zip")
        try:
            with self.db.lock:
                create_data_backup(
                    self.db.conn, self.settings_file, archive_path
                )
        except Exception as error:
            QMessageBox.critical(
                self,
                t("msg_error"),
                t("settings_export_error").format(error),
            )
            return

        QMessageBox.information(
            self,
            t("msg_success"),
            t("settings_exported").format(archive_path),
        )

    def import_data_backup(self, dialog: SettingsDialog):
        filename, _ = QFileDialog.getOpenFileName(
            self,
            t("settings_import_title"),
            str(self.user_data_dir),
            t("settings_zip_filter"),
        )
        if not filename:
            return

        reply = QMessageBox.question(
            self,
            t("settings_import_title"),
            t("settings_import_confirm"),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        self.save_settings()
        database_path = self.user_data_dir / "bookmarks.db"
        restore_error = None
        try:
            with self.db.lock:
                restore_data_backup(
                    Path(filename), self.user_data_dir, self.db.conn
                )
        except Exception as error:
            restore_error = error

        try:
            self.db = Database(database_path)
        except Exception as error:
            QMessageBox.critical(
                self,
                t("msg_error"),
                t("settings_import_error").format(
                    f"{restore_error}\n{error}" if restore_error else error
                ),
            )
            return

        if restore_error:
            QMessageBox.critical(
                self,
                t("msg_error"),
                t("settings_import_error").format(restore_error),
            )
            return

        self.load_settings()
        self.metinleri_guncelle()
        dialog.set_library_path(self.library_path)
        QMessageBox.information(self, t("msg_success"), t("settings_imported"))

    def load_pdf_bookmarks(self):
        pdf_root_name = t("pdf_bookmarks_root")
        from lang import get_all_translations_for_key
        possible_roots = get_all_translations_for_key("pdf_bookmarks_root")

        for i in range(self.tree.topLevelItemCount()):
            item = self.tree.topLevelItem(i)
            if item.text(0) in possible_roots:
                self.tree.takeTopLevelItem(i)
                break

        viewer = self.get_current_viewer()
        if not viewer or not viewer.doc:
            return

        toc = viewer.doc.get_toc()
        self.pdf_root = QTreeWidgetItem([pdf_root_name])
        self.tree.addTopLevelItem(self.pdf_root)
        parents = {0: self.pdf_root}

        for level, title, page in toc:
            item = QTreeWidgetItem([title])
            # Markdown ise page string (anchor id), PDF ise int (sayfa no)
            target_val = page if isinstance(page, str) else (page - 1)
            item.setData(0, Qt.UserRole + 100, target_val)
            parent_item = parents.get(level - 1, self.pdf_root)
            parent_item.addChild(item)
            parents[level] = item

        self.pdf_root.setExpanded(True)

    def zoom_in(self):
        viewer = self.get_current_viewer()
        if viewer:
            viewer.zoom_in()
            self.update_zoom_label()

    def zoom_out(self):
        viewer = self.get_current_viewer()
        if viewer:
            viewer.zoom_out()
            self.update_zoom_label()

    def zoom_100(self):
        viewer = self.get_current_viewer()
        if viewer:
            viewer.zoom_100()
            self.update_zoom_label()

    def zoom_fit(self):
        viewer = self.get_current_viewer()
        if viewer:
            viewer.fit_to_width()
            self.update_zoom_label()

    def update_zoom_label(self):
        viewer = self.get_current_viewer()
        if viewer:
            self.zoom_label.setText(t("lbl_zoom").format(int(viewer.zoom * 100)))
        else:
            self.zoom_label.setText("")

    def next_page(self):
        viewer = self.get_current_viewer()
        if viewer:
            viewer.next_page()
            self.update_page_label()

    def previous_page(self):
        viewer = self.get_current_viewer()
        if viewer:
            viewer.previous_page()
            self.update_page_label()

    def update_page_label(self):
        viewer = self.get_current_viewer()
        if not viewer or not viewer.doc:
            self.page_label.setText(t("lbl_page_empty"))
            self.page_edit.clear()
            return
        if self.sender() and isinstance(self.sender(), DOCViewer) and self.sender() != viewer:
            return
        current = viewer.current_page + 1
        total = len(viewer.doc)
        self.page_label.setText(t("lbl_page").format(current, total))
        self.page_edit.setText(str(current))

    def add_group(self):
        name, ok = QInputDialog.getText(self, t("dlg_new_group"), t("dlg_group_name"))
        if ok and name.strip():
            self.db.add_group(name.strip())
            self.load_groups()

    def load_groups(self):
        self.tree.clear()
        groups = self.db.get_groups()
        nodes = {}
        for gid, name, parent_id in groups:
            item = QTreeWidgetItem([name])
            item.setData(0, Qt.UserRole, gid)
            nodes[gid] = item
        for gid, name, parent_id in groups:
            item = nodes[gid]
            if parent_id and parent_id in nodes:
                nodes[parent_id].addChild(item)
            else:
                self.tree.addTopLevelItem(item)
        bookmarks = self.db.get_bookmarks()
        for bookmark in bookmarks:
            bid, title, pdf_path, page, group_id = bookmark[0], bookmark[1], bookmark[2], bookmark[3], bookmark[4]
            if group_id in nodes:
                bookmark_item = QTreeWidgetItem([title])
                bookmark_item.setData(0, Qt.UserRole + 1, bid)
                nodes[group_id].addChild(bookmark_item)
        self.load_pdf_bookmarks()
    def toggle_bookmarks(self):
        sizes = self.splitter.sizes()
        if sizes[0] > 0:
            self._last_left_width = sizes[0]
            sizes[0] = 0
        else:
            sizes[0] = getattr(self, "_last_left_width", 250)
        self.splitter.setSizes(sizes)

    def save_settings(self):
        sizes = self.splitter.sizes()
        is_visible = sizes[0] > 0
        if not is_visible and hasattr(self, "_last_left_width"):
            sizes[0] = self._last_left_width
        data = {
            "window_width": self.width(),
            "window_height": self.height(),
            "splitter_sizes": sizes,
            "bookmarks_visible": is_visible,
            "pdf_history": getattr(self, "pdf_history", []),
            "library_path": str(self.library_path),
            "last_opened_dir": self.last_opened_dir,
            "last_active_index": self.right_stack.currentIndex(),
            "language": "tr" if self.combo_lang.currentIndex() == 0 else "en",
            "theme": self.current_theme
        }
        with open(self.settings_file, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)

    def load_settings(self):
        if not self.settings_file.exists():
            self.update_dashboard_ui([])
            self.apply_theme()
            return
        try:
            with open(self.settings_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            set_lang(data.get("language", "tr"))
            self.combo_lang.blockSignals(True)
            self.combo_lang.setCurrentIndex(0 if data.get("language") == "tr" else 1)
            self.combo_lang.blockSignals(False)
            self.resize(data.get("window_width", 1400), data.get("window_height", 900))
            self.pdf_history = data.get("pdf_history", [])
            configured_library_path = data.get("library_path")
            if isinstance(configured_library_path, str) and configured_library_path:
                self.library_path = Path(configured_library_path).expanduser()
            else:
                self.library_path = self.default_library_path
            last_opened_dir = data.get("last_opened_dir")
            self.last_opened_dir = (
                last_opened_dir if isinstance(last_opened_dir, str) else None
            )
            self.current_theme = data.get("theme", "light")
            self.apply_theme()
            self.update_dashboard_ui(self.pdf_history)
        except Exception:
            self.update_dashboard_ui([])
            self.apply_theme()

    def update_dashboard_ui(self, history):
        while self.grid_layout.count():
            item = self.grid_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        try:
            self.stats_label.setText(t("dash_stats").format(self.db.get_bookmark_count(), self.db.get_group_count()))
        except Exception:
            pass

        if not history:
            lbl = QLabel(t("dash_empty"))
            lbl.setStyleSheet("color: #7f8c8d; font-style: italic;")
            self.grid_layout.addWidget(lbl, 0, 0)
            return

        for idx, d in enumerate(history):
            if not isinstance(d, dict):
                continue
            card = RecentFileCard(d, theme=self.current_theme)
            card.clicked.connect(self.open_from_dashboard_card)
            card.delete_requested.connect(self.delete_from_dashboard_card)
            self.grid_layout.addWidget(card, idx // 4, idx % 4)
    def open_from_dashboard_card(self, d):
        if d and Path(d.get("path", "")).exists():
            self.open_in_tab(d["path"], d.get("page", 0))

    def delete_from_dashboard_card(self, d):
        self.pdf_history = [i for i in self.pdf_history if i.get("path") != d.get("path")]
        self.update_dashboard_ui(self.pdf_history)
        self.save_settings()

    def closeEvent(self, event):
        workers = (
            self.pdf_search_workers
            + self.global_search_workers
            + [
                worker for worker in (
                    getattr(self, "office_worker", None),
                    getattr(self, "print_worker", None),
                    getattr(self.export_manager, "table_worker", None),
                )
                if worker is not None
            ]
        )
        for index in range(self.tab_widget.count()):
            viewer = self.tab_widget.widget(index)
            for worker in (
                getattr(viewer, "_render_worker", None),
                getattr(viewer, "ocr_thread", None),
            ):
                if worker is not None:
                    workers.append(worker)

        workers = list({id(worker): worker for worker in workers}.values())
        for worker in workers:
            worker.requestInterruption()
        for worker in workers:
            worker.wait()
        self.save_settings()
        self.db.conn.close()
        super().closeEvent(event)

    def _clean_md_line_for_search(self, raw_line, fallback_query):
        # [Link Metni](url) -> Link Metni
        cleaned = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', raw_line)
        # **kalın**, *italik*, `kod`, # başlık, > alıntı işaretlerini temizle
        cleaned = re.sub(r'[#*`>_~]+', '', cleaned).strip()
        # Çok uzunsa ilk 60 karakteri al (satır içi kırılmalardan etkilenmemek için)
        return cleaned[:60].strip() if len(cleaned) >= 4 else fallback_query


    def search_in_pdf(self):
        text_val = self.pdf_search.text().strip()
        if not text_val:
            return

        request_id = self._begin_search_request()
        search_scope = self.search_mode.currentIndex()
        if search_scope == 1:
            self.search_dashboard_and_tabs(text_val, request_id)
            return
        if search_scope == 2:
            self.search_configured_library(text_val, request_id)
            return

        viewer = self.get_current_viewer()
        if not viewer or not viewer.current_file_path:
            self.btn_search.setEnabled(True)
            return

        # Açık sekme MarkdownViewer ise PyMuPDF yerine yerleşik web aramasını kullan
        if isinstance(viewer, MarkdownViewer):
            self.search_results.clear()
            try:
                content = Path(viewer.current_file_path).read_text(encoding="utf-8")
                lines = content.splitlines()
                matches = [
                    (idx + 1, line.strip())
                    for idx, line in enumerate(lines)
                    if text_val.lower() in line.lower()
                ]
                self.search_results.addItem(f"Toplam Eşleşme: {len(matches)}")
                for occ_idx, (line_no, line_text) in enumerate(matches):
                    snippet = line_text[:65] + ("..." if len(line_text) > 65 else "")
                    itm = QListWidgetItem(f"Satır {line_no}: {snippet}")
                    itm.setData(Qt.UserRole, ("md_line", occ_idx, text_val))
                    self.search_results.addItem(itm)

                if matches:
                    viewer.jump_to_occurrence(text_val, 0)
            except Exception as e:
                self.search_results.addItem(f"Arama hatası: {e}")
            self.btn_search.setEnabled(True)
            return

        self.search_results.clear()
        self.search_results.addItem("Aranıyor...")
        self.btn_search.setEnabled(False)

        self.search_worker = PDFSearchWorker(viewer.current_file_path, text_val)
        worker = self.search_worker
        worker.results_ready.connect(
            lambda pages, rid=request_id: self._on_search_finished(rid, pages)
        )
        worker.error.connect(
            lambda message, rid=request_id: self._on_search_error(rid, message)
        )
        worker.finished.connect(
            lambda worker=worker: self._release_pdf_search_worker(worker)
        )
        self.pdf_search_workers.append(worker)
        self.search_worker.start()

    def _begin_search_request(self):
        self.search_request_id += 1
        for worker in self.pdf_search_workers + self.global_search_workers:
            worker.requestInterruption()
        return self.search_request_id

    def _start_global_search(
        self, query, request_id, bookmarks, groups, document_paths,
        library_path=None,
    ):
        self.search_results.clear()
        self.search_results.addItem(t("search_wait"))
        self.btn_search.setEnabled(False)

        worker = GlobalLibrarySearchWorker(
            request_id,
            query,
            bookmarks,
            groups,
            document_paths,
            library_path,
        )
        worker.results_ready.connect(self._on_global_search_finished)
        worker.error.connect(self._on_global_search_error)
        worker.finished.connect(
            lambda worker=worker: self._release_global_search_worker(worker)
        )
        self.global_search_workers.append(worker)
        worker.start()

    def search_dashboard_and_tabs(self, query, request_id):
        document_paths = set()
        for item in getattr(self, "pdf_history", []):
            path = item.get("path", "") if isinstance(item, dict) else item
            if path:
                document_paths.add(str(path))

        for index in range(self.tab_widget.count()):
            viewer = self.tab_widget.widget(index)
            for path in (
                getattr(viewer, "current_file_path", None),
                getattr(viewer, "original_file_path", None),
            ):
                if path:
                    document_paths.add(str(path))

        bookmarks = self.db.get_bookmarks()
        scoped_bookmarks = [
            bookmark
            for bookmark in bookmarks
            if bookmark[2] and str(bookmark[2]) in document_paths
        ]
        groups = dict((gid, name) for gid, name, _ in self.db.get_groups())
        self._start_global_search(
            query, request_id, scoped_bookmarks, groups, document_paths
        )

    def search_configured_library(self, query, request_id):
        library_root = Path(self.library_path).expanduser()
        bookmarks = []
        if library_root.is_dir():
            resolved_root = library_root.resolve()
            bookmarks = [
                bookmark
                for bookmark in self.db.get_bookmarks()
                if bookmark[2]
                and self._is_path_within(Path(bookmark[2]), resolved_root)
            ]
        groups = dict((gid, name) for gid, name, _ in self.db.get_groups())
        self._start_global_search(
            query,
            request_id,
            bookmarks,
            groups,
            set(),
            library_root,
        )

    @staticmethod
    def _is_path_within(path, root):
        try:
            path.resolve().relative_to(root)
            return True
        except (OSError, ValueError):
            return False

    def _release_global_search_worker(self, worker):
        if worker in self.global_search_workers:
            self.global_search_workers.remove(worker)

    def _release_pdf_search_worker(self, worker):
        if worker in self.pdf_search_workers:
            self.pdf_search_workers.remove(worker)

    def _on_global_search_finished(self, request_id, matches, errors):
        if request_id != self.search_request_id:
            return

        self.btn_search.setEnabled(True)
        self.search_results.clear()
        self.search_results.addItem(
            t("search_matches_count").format(len(matches))
        )
        for payload, preview in matches:
            itm = QListWidgetItem(preview)
            itm.setData(Qt.UserRole, payload)
            self.search_results.addItem(itm)

        if errors:
            error_item = QListWidgetItem(
                t("search_errors_count").format(len(errors))
            )
            error_item.setToolTip("\n".join(errors))
            self.search_results.addItem(error_item)

    def _on_global_search_error(self, request_id, err_msg):
        if request_id != self.search_request_id:
            return

        self.btn_search.setEnabled(True)
        self.search_results.clear()
        self.search_results.addItem(f"Arama hatası: {err_msg}")


    def _on_search_finished(self, request_id, pages):
        if request_id != self.search_request_id:
            return
        self.btn_search.setEnabled(True)
        self.search_results.clear()
        self.search_results.addItem(f"Toplam Eşleşme: {len(pages)}")
        for p, r in pages:
            itm = QListWidgetItem(f"Sayfa {p + 1}")
            itm.setData(Qt.UserRole, (p, r))
            self.search_results.addItem(itm)


    def _on_search_error(self, request_id, err_msg):
        if request_id != self.search_request_id:
            return
        self.btn_search.setEnabled(True)
        self.search_results.clear()
        self.search_results.addItem(f"Arama hatası: {err_msg}")

    def goto_search_result(self, item):
        data = item.data(Qt.UserRole)
        if not data:
            return

        # 1. Açık Markdown belgesindeki arama sonucuna çift tıklandıysa:
        if isinstance(data, tuple) and data[0] == "md_line":
            _, occ_idx, query_str = data
            v = self.get_current_viewer()
            if isinstance(v, MarkdownViewer):
                self.record_navigation_point()
                v.jump_to_occurrence(query_str, occ_idx)
            return

        # 2. Tüm kütüphane aramasındaki bir Markdown dosyasına çift tıklandıysa:
        if isinstance(data, tuple) and data[0] == "global_md":
            _, md_path, occ_idx, query_str = data
            if Path(md_path).exists():
                self.record_navigation_point()
                self.open_in_tab(md_path)
                v = self.get_current_viewer()
                if isinstance(v, MarkdownViewer):
                    v.jump_to_occurrence(query_str, occ_idx)
            return

        if isinstance(data, tuple) and data[0] == "global_bookmark":
            _, bid = data
            bookmark = self.db.get_bookmark(bid)
            if not bookmark:
                return

            self.record_navigation_point()
            self.active_bookmark_id = bid
            _, _, pdf_path, page, _ = bookmark[:5]
            x, y, current_note = 0, 0, ""
            if len(bookmark) > 7:
                _, _, _, _, _, bx, by, bnote = bookmark[:8]
                x, y, current_note = bx or 0, by or 0, bnote or ""

            if Path(pdf_path).exists():
                self.open_in_tab(pdf_path, page, x, y)
                self.note_edit.setReadOnly(True)
                html_note = current_note
                for full_match, link_content in re.findall(r'(\[\[([^\]]+)\]\])', current_note):
                    html_link = f'<a href="{link_content}" style="color: #0078ff; font-weight: bold; text-decoration: none;">{full_match}</a>'
                    html_note = html_note.replace(full_match, html_link)
                self.note_edit.setHtml(html_note.replace("\n", "<br>"))
                self.note_edit.setEnabled(True)
                self.btn_save_note.setEnabled(True)
            else:
                QMessageBox.warning(self, t("msg_error"), t("err_file_relocate").format(pdf_path))
            return

        if isinstance(data, tuple) and data[0] == "global_pdf":
            _, pdf_path, page_no, rects = data
            if Path(pdf_path).exists():
                self.record_navigation_point()
                self.open_in_tab(pdf_path, page_no)
                v = self.get_current_viewer()
                if v:
                    v.set_highlight([(page_no, r) for r in rects])
                    v.render_page()
            return

        v = self.get_current_viewer()
        if v:
            self.record_navigation_point()
            p, r = data
            v.goto_page(p)
            v.set_highlight([(p, r)])
            v.render_page()

    def filter_tree(self, text):
        for i in range(self.tree.topLevelItemCount()):
            self.filter_item(self.tree.topLevelItem(i), text.lower().strip())

    def filter_item(self, item, text):
        vis = text in item.text(0).lower()
        c_vis = any(self.filter_item(item.child(i), text) for i in range(item.childCount()))
        item.setHidden(not (vis or c_vis or text == ""))
        return vis or c_vis

    def show_tree_menu(self, pos):
        item = self.tree.itemAt(pos)
        if not item:
            m = QMenu()
            a_new_gr = m.addAction(t("btn_group"))
            action = m.exec(self.tree.mapToGlobal(pos))
            if action == a_new_gr:
                self.add_group()
            return

        m = QMenu()
        group_id = item.data(0, Qt.UserRole)
        bookmark_id = item.data(0, Qt.UserRole + 1)
        pdf_page = item.data(0, Qt.UserRole + 100)

        if pdf_page is not None:
            return

        if group_id is not None:
            a_bm = m.addAction(t("menu_add_bookmark"))
            r_gr = m.addAction(t("dlg_rename_group"))
            d_gr = m.addAction(t("msg_delete_group_confirm").format(item.text(0)))

            action = m.exec(self.tree.mapToGlobal(pos))

            if action == a_bm:
                self.add_bookmark(item)
            elif action == r_gr:
                self.rename_group_action(item)
            elif action == d_gr:
                self.delete_group_action(item)

        elif bookmark_id is not None:
            r_bm = m.addAction(t("menu_rename_bookmark"))
            l_bm = m.addAction("🔗 Link Ekle (Çapraz PDF)")
            d_bm = m.addAction(t("menu_del_bookmark"))

            action = m.exec(self.tree.mapToGlobal(pos))

            if action == r_bm:
                self.rename_bookmark(item)
            elif action == l_bm:
                self.add_cross_pdf_link(item)
            elif action == d_bm:
                self.delete_bookmark(item)

    def add_bookmark(self, item):
        viewer = self.get_current_viewer()
        if not viewer or not item:
            return
        current_page = viewer.current_page
        group_id = item.data(0, Qt.UserRole)
        title, ok = QInputDialog.getText(self, t("dlg_new_bookmark"), t("dlg_bookmark_name"))
        if not ok or not title.strip():
            return

        kaydedilecek_yol = getattr(viewer, "original_file_path", viewer.current_file_path)
        self.db.add_bookmark(title.strip(), kaydedilecek_yol, current_page, group_id)
        self.load_groups()
        viewer.goto_page(current_page)
        self.update_page_label()

    def delete_bookmark(self, item):
        if not item:
            return
        bookmark_id = item.data(0, Qt.UserRole + 1)
        if not bookmark_id:
            return
        self.db.delete_bookmark(bookmark_id)
        self.load_groups()

    def rename_bookmark(self, item):
        if not item:
            return
        bookmark_id = item.data(0, Qt.UserRole + 1)
        if not bookmark_id:
            return
        name, ok = QInputDialog.getText(self, t("dlg_rename_bookmark"), t("dlg_new_name"), text=item.text(0))
        if ok and name.strip():
            self.db.rename_bookmark(bookmark_id, name.strip())
            self.load_groups()

    def open_bookmark(self, item):
        self.record_navigation_point()
        pdf_page = item.data(0, Qt.UserRole + 100)

        if pdf_page is not None:
            viewer = self.get_current_viewer()
            if viewer:
                if isinstance(pdf_page, str) and hasattr(viewer, "scroll_to_anchor"):
                    viewer.scroll_to_anchor(pdf_page, item.text(0))
                else:
                    viewer.goto_page(pdf_page)
                self.update_page_label()
                self.right_stack.setCurrentIndex(1)
            return


        bookmark_id = item.data(0, Qt.UserRole + 1)
        if not bookmark_id:
            self.active_bookmark_id = None
            self.note_edit.clear()
            self.note_edit.setEnabled(False)
            self.btn_save_note.setEnabled(False)
            return

        bookmark = self.db.get_bookmark(bookmark_id)
        if not bookmark:
            return

        self.active_bookmark_id = bookmark_id
        pdf_path = bookmark[2]
        page = bookmark[3]
        x = bookmark[5] if len(bookmark) > 5 else 0
        y = bookmark[6] if len(bookmark) > 6 else 0
        current_note = bookmark[7] if len(bookmark) > 7 and bookmark[7] else ""

        if "[[" in current_note and "]]" in current_note and "#" in current_note:
            try:
                link_match = re.search(r'\[\[([^\]]+)\]\]', current_note)
                if link_match:
                    _, page_str = link_match.group(1).split("#", 1)
                    clean_page_str = "".join(filter(str.isdigit, page_str))
                    if clean_page_str:
                        page = int(clean_page_str) - 1
            except Exception as e:
                print("Link çözümleme hatası:", e)

        path_check = Path(pdf_path)
        if not path_check.exists():
            reply = QMessageBox.question(
                self, t("msg_error"), t("err_file_relocate").format(pdf_path),
                QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes
            )
            if reply == QMessageBox.Yes:
                new_filename, _ = QFileDialog.getOpenFileName(
                    self, t("dlg_file_relocate"),
                    QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation),
                    "PDF Files (*.pdf)"
                )
                if not new_filename:
                    return
                self.db.update_pdf_path(pdf_path, new_filename)
                pdf_path = new_filename
            else:
                return
        self.open_in_tab(pdf_path, page, x, y)

        self.note_edit.setReadOnly(True)
        html_note = current_note
        matches = re.findall(r'(\[\[([^\]]+)\]\])', current_note)
        for full_match, link_content in matches:
            html_link = f'<a href="{link_content}" style="color: #0078ff; font-weight: bold; text-decoration: none;">{full_match}</a>'
            html_note = html_note.replace(full_match, html_link)

        self.note_edit.setHtml(html_note.replace("\n", "<br>"))
        self.note_edit.setEnabled(True)
        self.btn_save_note.setEnabled(True)

    def rename_group_action(self, item):
        if not item:
            return
        group_id = item.data(0, Qt.UserRole)
        if group_id is None:
            return

        name, ok = QInputDialog.getText(
            self,
            t("dlg_rename_group"),
            t("dlg_group_new_name"),
            text=item.text(0)
        )
        if ok and name.strip():
            self.db.rename_group(group_id, name.strip())
            self.load_groups()

    def delete_group_action(self, item):
        if not item:
            return
        group_id = item.data(0, Qt.UserRole)
        if group_id is None:
            return
        msg_text = t("msg_delete_group_confirm").format(item.text(0))
        reply = QMessageBox.question(self, t("msg_warning"), msg_text, QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.db.delete_group(group_id)
            self.load_groups()

    def enable_note_editing(self):
        if not self.active_bookmark_id or not self.note_edit.isReadOnly():
            return
        bookmark_data = self.db.get_bookmark(self.active_bookmark_id)
        if bookmark_data and len(bookmark_data) > 7:
            current_note = bookmark_data[7] or ""
            self.note_edit.setReadOnly(False)
            self.note_edit.setTextInteractionFlags(Qt.TextEditable | Qt.TextSelectableByMouse)
            self.note_edit.setPlainText(current_note)

    def save_current_bookmark_note(self):
        if not self.active_bookmark_id:
            return
        try:
            notes_text = self.note_edit.toPlainText().strip()
            self.db.update_bookmark_notes(self.active_bookmark_id, notes_text)
            self.note_edit.setReadOnly(True)
            self.note_edit.setTextInteractionFlags(Qt.LinksAccessibleByMouse)

            html_note = notes_text
            matches = re.findall(r'(\[\[([^\]]+)\]\])', notes_text)
            for full_match, link_content in matches:
                html_link = f'<a href="{link_content}" style="color: #0078ff; font-weight: bold; text-decoration: none;">{full_match}</a>'
                html_note = html_note.replace(full_match, html_link)
            self.note_edit.setHtml(html_note.replace("\n", "<br>"))
        except Exception as e:
            QMessageBox.warning(self, t("msg_error"), t("err_save_note").format(e))

    def handle_knowledge_link(self, url):
        try:
            self.record_navigation_point()
            raw_str = url.toString().replace("link://", "").strip().rstrip("/")
            if not raw_str and url.isValid():
                raw_str = url.toDisplayString().replace("link://", "").strip().rstrip("/")
            if not raw_str:
                return

            if "#" not in raw_str:
                target_file, target_page = raw_str, 0
            else:
                target_file, page_str = raw_str.split("#", 1)
                target_file = target_file.strip()
                clean_page_str = "".join(filter(str.isdigit, page_str))
                target_page = int(clean_page_str) - 1 if clean_page_str else 0

            viewer = self.get_current_viewer()
            if viewer and (Path(viewer.current_file_path).name == target_file or viewer.current_file_path == target_file):
                viewer.goto_page(target_page)
                self.update_page_label()
                self.right_stack.setCurrentIndex(1)
                return

            full_pdf_path = None
            if viewer:
                sibling_path = Path(viewer.current_file_path).parent / target_file
                if sibling_path.exists() and sibling_path.is_file():
                    full_pdf_path = str(sibling_path)

            if not full_pdf_path:
                for item in getattr(self, "pdf_history", []):
                    p_str = item.get("path", "") if isinstance(item, dict) else item
                    if p_str and (p_str.endswith(target_file) or Path(p_str).name == target_file):
                        if Path(p_str).exists() and Path(p_str).is_file():
                            full_pdf_path = p_str
                            break

            if not full_pdf_path:
                QMessageBox.warning(self, t("msg_error"), t("err_target_pdf").format(target_file))
                return

            self.open_in_tab(full_pdf_path, target_page)
        except Exception as e:
            print("Knowledge link hatası:", e)

    def add_bookmark_from_selection(self, page, text, x, y):
        viewer = self.sender()
        if not viewer:
            return
        if self.selected_group_id is None:
            QMessageBox.warning(self, t("msg_warning"), t("warn_select_group"))
            return
        kaydedilecek_yol = getattr(viewer, "original_file_path", viewer.current_file_path)
        self.db.add_bookmark(text, kaydedilecek_yol, page, self.selected_group_id, x, y)
        self.load_groups()
        if not isinstance(viewer, MarkdownViewer):
            viewer.goto_page(page)
        self.update_page_label()


    def _save_and_reload_pdf(self, viewer):
        viewer.wait_for_render()
        try:
            if getattr(viewer.doc, "is_repaired", False):
                temp_path = viewer.current_file_path + ".tmp.pdf"
                viewer.doc.save(temp_path, garbage=4, deflate=True)
                viewer.doc.close()
                os.replace(temp_path, viewer.current_file_path)
                viewer.doc = pymupdf.open(viewer.current_file_path)
            else:
                viewer.doc.saveIncr()

            if hasattr(viewer.doc, "reload"):
                viewer.doc.reload()
        except Exception as e:
            QMessageBox.warning(self, t("msg_error"), t("err_save_failed").format(e))
        finally:
            if viewer.doc and not viewer.doc.is_closed:
                viewer.render_page()

    def create_highlight(self, page_no, x0, y0, x1, y1):
        viewer = self.sender()
        if not viewer or not viewer.doc:
            return
        try:
            page = viewer.doc[page_no]
            rect = pymupdf.Rect(x0, y0, x1, y1)
            annot = page.add_rect_annot(rect)
            annot.set_colors(stroke=(1, 1, 0), fill=(1, 1, 0))
            annot.set_opacity(0.4)
            annot.update()
            self._save_and_reload_pdf(viewer)
        except Exception as e:
            QMessageBox.warning(self, t("msg_error"), t("err_highlight_readonly").format(e))

    def delete_highlight(self, page_no, x, y):
        viewer = self.sender()
        if not viewer or not viewer.doc:
            return
        try:
            page = viewer.doc[page_no]
            point = pymupdf.Point(x, y)
            annot = page.first_annot
            while annot:
                next_annot = annot.next
                if annot.rect.contains(point):
                    page.delete_annot(annot)
                    self._save_and_reload_pdf(viewer)
                    return
                annot = next_annot
        except Exception as e:
            QMessageBox.warning(self, t("msg_error"), t("err_delete_failed").format(e))

    def add_page_note(self, page_no, x, y):
        viewer = self.sender()
        if not viewer or not viewer.doc:
            return

        text, ok = QInputDialog.getMultiLineText(self, t("dlg_add_note_title"), t("dlg_add_note_msg"))
        if not ok or not text.strip():
            return

        try:
            page = viewer.doc[page_no]
            lines = text.strip().split('\n')
            max_len = max(*(len(line) for line in lines), 1)

            total_lines = sum([max(1, len(line) // 35) for line in lines])
            width = min(250, max(100, max_len * 7))
            height = max(40, total_lines * 16 + 15)
            rect = pymupdf.Rect(x, y, x + width, y + height)
            rect.normalize()
            annot = page.add_freetext_annot(
                rect, text.strip(), fontsize=11, fontname="helv",
                text_color=(1, 0, 0), fill_color=(1, 1, 0.8)
            )
            annot.set_border(width=1)
            annot.update()
            self._save_and_reload_pdf(viewer)
        except Exception as e:
            QMessageBox.warning(self, t("msg_error"), t("err_add_note_failed").format(e))
    def print_pdf(self):
        viewer = self.get_current_viewer()
        if not viewer or not viewer.doc:
            return
        printer = QPrinter(QPrinter.HighResolution)
        dialog = QPrintDialog(printer, self)
        dialog.setWindowTitle(t("print_title"))
        if dialog.exec() == QPrintDialog.Accepted:
            from_page = printer.fromPage() - 1
            to_page = printer.toPage() - 1
            if from_page < 0:
                from_page = 0
            if to_page < 0 or to_page >= len(viewer.doc):
                to_page = len(viewer.doc) - 1
            self.progress_dialog = QProgressDialog(t("prog_print_wait"), None, 0, 0, self)
            self.progress_dialog.setWindowTitle(t("print_title"))
            self.progress_dialog.setWindowModality(Qt.WindowModal)
            self.progress_dialog.show()
            self.print_worker = PDFPrintWorker(viewer.current_file_path, printer, from_page, to_page)
            self.print_worker.printed.connect(self._on_print_finished)
            self.print_worker.error.connect(self._on_print_error)
            self.print_worker.start()

    def _on_print_finished(self, printer_name):
        if hasattr(self, "progress_dialog"):
            self.progress_dialog.close()
        QMessageBox.information(self, t("print_title"), t("print_all_success").format(printer_name))

    def _on_print_error(self, err_msg):
        if hasattr(self, "progress_dialog"):
            self.progress_dialog.close()
        QMessageBox.warning(self, t("print_err"), err_msg)

    def goto_page(self):
        viewer = self.get_current_viewer()
        if not viewer or not viewer.doc:
            return
        try:
            page_no = int(self.page_edit.text())
            if 1 <= page_no <= viewer.doc.page_count:
                self.record_navigation_point()
                viewer.current_page = page_no - 1
                viewer.render_page()

        except ValueError:
            pass
        finally:
            self.update_page_label()

    def add_cross_pdf_link(self, item):
        bid = item.data(0, Qt.UserRole + 1)
        if not bid:
            return
        pdf_options = [t("link_wiz_newfile")]
        history = getattr(self, "pdf_history", [])
        pdf_options.extend([Path(d.get("path")).name for d in history if isinstance(d, dict) and d.get("path")])
        target_file, ok1 = QInputDialog.getItem(self, t("link_wiz_title"), t("link_wiz_q1"), pdf_options, 0, False)
        if not ok1 or not target_file:
            return
        final_file_name = target_file
        if target_file == t("link_wiz_newfile"):
            filename, _ = QFileDialog.getOpenFileName(
                self, t("link_wiz_open"),
                QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation),
                "PDF Files (*.pdf)"
            )
            if not filename:
                return
            final_file_name = Path(filename).name
            self.add_to_history(filename, 0)
        target_page, ok2 = QInputDialog.getInt(self, t("link_wiz_title"), t("link_wiz_q2").format(final_file_name), 1, 1, 9999, 1)
        if not ok2:
            return
        try:
            self.note_edit.setReadOnly(False)
            updated_text = self.note_edit.toPlainText() + f" [[{final_file_name}#{target_page}]]"
            self.db.update_bookmark_notes(bid, updated_text)
            self.note_edit.setReadOnly(True)
            html_note = updated_text
            for full_match, link_content in re.findall(r'(\[\[([^\]]+)\]\])', updated_text):
                html_link = f'<a href="{link_content}" style="color: #0078ff; font-weight: bold; text-decoration: none;">{full_match}</a>'
                html_note = html_note.replace(full_match, html_link)
            self.note_edit.setHtml(html_note.replace("\n", "<br>"))
        except Exception as e:
            QMessageBox.warning(self, t("msg_error"), str(e))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setApplicationName("Doc_Network")
    app.setDesktopFileName("Doc_Network")

    window = MainWindow()

    # DÜZELTİLMİŞ KISIM: sys.argv[1] kullanıldı
    if len(sys.argv) > 1:
        file_path = Path(sys.argv[1]).resolve()
        if file_path.exists() and file_path.is_file():
            window.open_in_tab(str(file_path))

    window.show()
    sys.exit(app.exec())
