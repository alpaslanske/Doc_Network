#!/usr/bin/env python3
from pathlib import Path
from PySide6.QtCore import Signal, Qt
from PySide6.QtWidgets import QFrame, QVBoxLayout, QHBoxLayout, QLabel, QPushButton


class RecentFileCard(QFrame):
    clicked = Signal(dict)
    delete_requested = Signal(dict)

    def __init__(self, file_data, theme="light", parent=None):
        super().__init__(parent)
        self.file_data = file_data
        self.theme = theme
        self.setup_ui()

    def setup_ui(self):
        self.setFrameShape(QFrame.StyledPanel)
        self.setObjectName("RecentCard")
        self.setCursor(Qt.PointingHandCursor)
        self.setFixedWidth(280)
        self.setMinimumHeight(90)

        if self.theme == "dark":
            self.setStyleSheet("""
                QFrame#RecentCard {
                    background-color: #1a2027;
                    border: 1px solid #2d3640;
                    border-radius: 12px;
                }
                QFrame#RecentCard:hover {
                    border: 1px solid #137886;
                }
            """)
        else:
            self.setStyleSheet("""
                QFrame#RecentCard {
                    background-color: #ffffff;
                    border: 1px solid #e2e8f0;
                    border-radius: 12px;
                }
                QFrame#RecentCard:hover {
                    background-color: #ffffff;
                    border: 1px solid #3498db;
                }
            """)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(8)

        header_layout = QHBoxLayout()
        path_str = self.file_data.get("path", "")
        suffix = Path(path_str).suffix.lower()
        icon_emoji = "📝" if suffix in [".docx", ".odt"] else "📘"

        lbl_icon = QLabel(icon_emoji)
        lbl_icon.setStyleSheet("font-size: 16px;")
        header_layout.addWidget(lbl_icon)

        file_name = Path(path_str).name
        lbl_name = QLabel(file_name)

        if self.theme == "dark":
            lbl_name.setStyleSheet("font-weight: bold; font-size: 13px; color: #d1d6dc;")
        else:
            lbl_name.setStyleSheet("font-weight: bold; font-size: 13px; color: #2c3e50;")

        metrics = lbl_name.fontMetrics()
        elided_text = metrics.elidedText(file_name, Qt.ElideRight, 180)
        lbl_name.setText(elided_text)
        header_layout.addWidget(lbl_name, 1)

        btn_delete = QPushButton("×")
        btn_delete.setFixedSize(20, 20)
        btn_delete.setCursor(Qt.ArrowCursor)
        if self.theme == "dark":
            btn_delete.setStyleSheet("""
                QPushButton { border: none; border-radius: 10px; color: #5a6673; font-size: 14px; font-weight: bold; background-color: transparent;}
                QPushButton:hover { background-color: #a93226; color: white; }
            """)
        else:
            btn_delete.setStyleSheet("""
                QPushButton { border: none; border-radius: 10px; color: #95a5a6; font-size: 14px; font-weight: bold; background-color: transparent;}
                QPushButton:hover { background-color: #e74c3c; color: white; }
            """)

        btn_delete.clicked.connect(self.emit_delete)
        header_layout.addWidget(btn_delete)
        main_layout.addLayout(header_layout)

        footer_layout = QHBoxLayout()
        page_no = self.file_data.get("page", 0) + 1
        lbl_page = QLabel(f" Sayfa {page_no}")

        if self.theme == "dark":
            lbl_page.setStyleSheet("""
                font-size: 11px;
                background-color: #137886;
                color: #ffffff;
                padding: 3px 8px;
                border-radius: 6px;
                font-weight: bold;
            """)
        else:
            lbl_page.setStyleSheet("""
                font-size: 11px;
                background-color: #e8f4fd;
                color: #2980b9;
                padding: 3px 8px;
                border-radius: 6px;
                font-weight: bold;
            """)

        footer_layout.addWidget(lbl_page)

        date_raw = self.file_data.get("date", "-")
        date_clean = date_raw.split(" ")[0] if " " in date_raw else date_raw
        lbl_date = QLabel(f"⏱️ {date_clean}")
        lbl_date.setStyleSheet("color: #95a5a6; font-size: 11px;")
        footer_layout.addWidget(lbl_date, 0, Qt.AlignRight)

        main_layout.addLayout(footer_layout)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.clicked.emit(self.file_data)
        super().mousePressEvent(event)

    def emit_delete(self):
        self.delete_requested.emit(self.file_data)
