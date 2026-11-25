import os
import sqlite3
import tempfile
import threading
from pathlib import Path
from typing import List, Tuple, Optional
from PySide6.QtCore import QStandardPaths

def get_user_data_dir() -> Path:
    data_dir = QStandardPaths.writableLocation(QStandardPaths.AppDataLocation)
    if not data_dir:
        raise RuntimeError("The application data directory is unavailable.")
    return Path(data_dir)


def migrate_legacy_database(source_path: Path, destination_path: Path) -> bool:
    if not source_path.is_file() or destination_path.exists():
        return False

    destination_path.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary_name = tempfile.mkstemp(
        prefix=f".{destination_path.name}.",
        suffix=".migrating",
        dir=destination_path.parent,
    )
    os.close(fd)
    temporary_path = Path(temporary_name)
    try:
        source = sqlite3.connect(source_path)
        try:
            destination = sqlite3.connect(temporary_path)
            try:
                source.backup(destination)
            finally:
                destination.close()
        finally:
            source.close()
        if destination_path.exists():
            return False
        os.replace(temporary_path, destination_path)
        return True
    finally:
        temporary_path.unlink(missing_ok=True)


class Database:
    def __init__(self, db_path: Optional[Path] = None) -> None:
        db_path = db_path or get_user_data_dir() / "bookmarks.db"
        db_path.parent.mkdir(parents=True, exist_ok=True)
        # Çoklu thread erişimlerinde kilitlenmeyi önlemek için kilit (lock) oluşturuyoruz
        self.lock = threading.Lock()

        self.conn = sqlite3.connect(db_path, check_same_thread=False)

        with self.lock:
            # SQLite'ı hızlandırmak ve kilitlenmeleri (database is locked) önlemek için WAL modu
            self.conn.execute("PRAGMA journal_mode=WAL;")
            self.conn.execute("PRAGMA synchronous=NORMAL;")

        self.create_tables()
        self.check_and_update_schema()

    def create_tables(self) -> None:
        with self.lock:
            self.conn.execute("""
            CREATE TABLE IF NOT EXISTS groups(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                parent_id INTEGER
            )
            """)

            self.conn.execute("""
            CREATE TABLE IF NOT EXISTS bookmarks(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL,
                pdf_path TEXT NOT NULL,
                page INTEGER NOT NULL,
                group_id INTEGER,
                x REAL DEFAULT 0,
                y REAL DEFAULT 0,
                notes TEXT
            )
            """)

            # Performans için indeksler
            self.conn.execute("CREATE INDEX IF NOT EXISTS idx_bookmarks_group ON bookmarks(group_id)")
            self.conn.execute("CREATE INDEX IF NOT EXISTS idx_bookmarks_pdf ON bookmarks(pdf_path)")

            self.conn.commit()

    def check_and_update_schema(self) -> None:
        with self.lock:
            try:
                cursor = self.conn.cursor()
                cursor.execute("PRAGMA table_info(bookmarks)")
                columns = [col[1] for col in cursor.fetchall()]

                if "notes" not in columns:
                    cursor.execute("ALTER TABLE bookmarks ADD COLUMN notes TEXT")
                if "x" not in columns:
                    cursor.execute("ALTER TABLE bookmarks ADD COLUMN x REAL DEFAULT 0")
                if "y" not in columns:
                    cursor.execute("ALTER TABLE bookmarks ADD COLUMN y REAL DEFAULT 0")
                self.conn.commit()
            except Exception as e:
                print("Veritabanı şema güncelleme hatası:", e)

    def add_group(self, name: str, parent_id: Optional[int] = None) -> None:
        with self.lock:
            self.conn.execute(
                "INSERT INTO groups(name, parent_id) VALUES(?,?)",
                (name, parent_id)
            )
            self.conn.commit()

    def get_groups(self) -> List[Tuple]:
        with self.lock:
            return self.conn.execute(
                "SELECT id, name, parent_id FROM groups ORDER BY name"
            ).fetchall()

    def get_group_count(self) -> int:
        with self.lock:
            res = self.conn.execute("SELECT COUNT(*) FROM groups").fetchone()
            return res[0] if res else 0

    def rename_group(self, group_id: int, new_name: str) -> None:
        with self.lock:
            self.conn.execute(
                "UPDATE groups SET name=? WHERE id=?",
                (new_name, group_id)
            )
            self.conn.commit()

    def delete_group(self, group_id: int) -> None:
        with self.lock:
            with self.conn:
                group = self.conn.execute(
                    "SELECT parent_id FROM groups WHERE id=?", (group_id,)
                ).fetchone()
                if group is None:
                    return

                self.conn.execute(
                    "UPDATE groups SET parent_id=? WHERE parent_id=?",
                    (group[0], group_id)
                )
                self.conn.execute(
                    "DELETE FROM bookmarks WHERE group_id=?", (group_id,)
                )
                self.conn.execute("DELETE FROM groups WHERE id=?", (group_id,))

    def add_bookmark(self, title: str, pdf_path: str, page: int, group_id: int, x: float = 0, y: float = 0) -> None:
        with self.lock:
            self.conn.execute(
                """
                INSERT INTO bookmarks(title, pdf_path, page, group_id, x, y)
                VALUES(?,?,?,?,?,?)
                """,
                (title, pdf_path, page, group_id, x, y)
            )
            self.conn.commit()

    def get_bookmarks(self) -> List[Tuple]:
        with self.lock:
            return self.conn.execute(
                """
                SELECT id, title, pdf_path, page, group_id, x, y, notes
                FROM bookmarks
                ORDER BY title
                """
            ).fetchall()

    def get_bookmark_count(self) -> int:
        with self.lock:
            res = self.conn.execute("SELECT COUNT(*) FROM bookmarks").fetchone()
            return res[0] if res else 0

    def get_bookmark(self, bookmark_id: int) -> Optional[Tuple]:
        with self.lock:
            return self.conn.execute(
                "SELECT id, title, pdf_path, page, group_id, x, y, notes FROM bookmarks WHERE id=?",
                (bookmark_id,)
            ).fetchone()

    def rename_bookmark(self, bookmark_id: int, new_title: str) -> None:
        with self.lock:
            self.conn.execute(
                "UPDATE bookmarks SET title=? WHERE id=?",
                (new_title, bookmark_id)
            )
            self.conn.commit()

    def delete_bookmark(self, bookmark_id: int) -> None:
        with self.lock:
            self.conn.execute("DELETE FROM bookmarks WHERE id=?", (bookmark_id,))
            self.conn.commit()

    def update_pdf_path(self, old_path: str, new_path: str) -> None:
        with self.lock:
            self.conn.execute(
                "UPDATE bookmarks SET pdf_path = ? WHERE pdf_path = ?",
                (new_path, old_path)
            )
            self.conn.commit()

    def update_bookmark_notes(self, bookmark_id: int, notes: str) -> None:
        with self.lock:
            self.conn.execute(
                "UPDATE bookmarks SET notes = ? WHERE id = ?",
                (notes, bookmark_id)
            )
            self.conn.commit()
