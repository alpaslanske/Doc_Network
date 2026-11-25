import os
import shutil
import subprocess
import sys
from pathlib import Path


SUPPORTED_MIME_TYPES = (
    "application/pdf;"
    "application/epub+zip;"
    "application/x-mobipocket-ebook;"
    "application/x-fictionbook+xml;"
    "application/vnd.comicbook+zip;"
    "text/markdown;"
    "text/x-markdown;"
    "text/plain;"
    "text/csv;"
    "application/rtf;"
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document;"
    "application/vnd.oasis.opendocument.text;"
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;"
    "application/vnd.oasis.opendocument.spreadsheet;"
    "application/vnd.openxmlformats-officedocument.presentationml.presentation;"
    "application/vnd.oasis.opendocument.presentation;"
)


def _resolve_exec_command(main_script_path: Path) -> str:
    if "APPIMAGE" in os.environ:
        return f'"{os.environ["APPIMAGE"]}" %f'
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}" %f'
    return f'"{sys.executable}" "{main_script_path.resolve()}" %f'


def register_context_menu(main_script_path: Path, icon_path: Path | None = None) -> Path:
    apps_dir = Path.home() / ".local" / "share" / "applications"
    apps_dir.mkdir(parents=True, exist_ok=True)

    desktop_file = apps_dir / "Doc_Network.desktop"
    exec_cmd = _resolve_exec_command(main_script_path)
    icon_value = str(icon_path.resolve()) if icon_path and icon_path.exists() else "application-pdf"

    content = (
        "[Desktop Entry]\n"
        "Type=Application\n"
        "Name=Doc_Network\n"
        "Comment=Belge Görüntüleyici ve Çalışma Alanı\n"
        f"Exec={exec_cmd}\n"
        f"Icon={icon_value}\n"
        "Terminal=false\n"
        "Categories=Office;Viewer;\n"
        f"MimeType={SUPPORTED_MIME_TYPES}\n"
    )

    desktop_file.write_text(content, encoding="utf-8")
    desktop_file.chmod(0o755)

    if shutil.which("update-desktop-database"):
        subprocess.run(
            ["update-desktop-database", str(apps_dir)],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
        )

    return desktop_file
