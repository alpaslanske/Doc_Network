from pathlib import Path

from PySide6.QtCore import Signal
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
)

from lang import t


class SettingsDialog(QDialog):
    backup_export_requested = Signal()
    backup_import_requested = Signal()

    def __init__(self, library_path: Path, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle(t("settings_title"))
        self.setMinimumWidth(520)

        self.library_path_input = QLineEdit(str(library_path))
        self.browse_button = QPushButton(t("settings_browse"))
        self.browse_button.clicked.connect(self.browse_library_path)

        path_row = QHBoxLayout()
        path_row.addWidget(self.library_path_input, 1)
        path_row.addWidget(self.browse_button)

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(t("settings_library_label")))
        layout.addLayout(path_row)

        backup_row = QHBoxLayout()
        self.export_button = QPushButton(t("settings_backup_export"))
        self.import_button = QPushButton(t("settings_backup_import"))
        self.export_button.clicked.connect(self.backup_export_requested)
        self.import_button.clicked.connect(self.backup_import_requested)
        backup_row.addWidget(self.export_button)
        backup_row.addWidget(self.import_button)
        layout.addLayout(backup_row)

        self.button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save
            | QDialogButtonBox.StandardButton.Cancel
        )
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)

        self._selected_library_path = library_path

    @property
    def selected_library_path(self) -> Path:
        return self._selected_library_path

    def set_library_path(self, library_path: Path) -> None:
        self._selected_library_path = library_path
        self.library_path_input.setText(str(library_path))

    def browse_library_path(self) -> None:
        current_path = Path(self.library_path_input.text()).expanduser()
        initial_path = str(current_path if current_path.is_dir() else Path.home())
        selected_path = QFileDialog.getExistingDirectory(
            self, t("settings_browse_title"), initial_path
        )
        if selected_path:
            self.library_path_input.setText(selected_path)

    def accept(self) -> None:
        selected_path = Path(self.library_path_input.text()).expanduser()
        if not selected_path.is_dir():
            QMessageBox.warning(
                self, t("msg_warning"), t("settings_library_invalid")
            )
            return

        self._selected_library_path = selected_path.resolve()
        super().accept()
