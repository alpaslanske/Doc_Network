import os
import sqlite3
import threading
from pathlib import Path
from typing import List, Tuple, Optional

def get_system_documents_dir() -> Path:
    # Eğer uygulama AppImage paketi olarak çalışıyorsa gerçek klasörü bulur,
    # Normal Python dosyası olarak çalışıyorsa bulunduğu klasörü bulur.
    if "APPIMAGE" in os.environ:
        return Path(os.environ["APPIMAGE"]).parent
    else:
        return Path(os.path.realpath(__file__)).parent

LocationForBookmarks = get_system_documents_dir()
DB_DIR = LocationForBookmarks
DB_PATH = DB_DIR / "data" / "bookmarks.db"

class Database:
    def __init__(self) -> None:
        DB_PATH.parent.mkdir(parents=True, exist_ok=True)
        # Çoklu thread erişimlerinde kilitlenmeyi önlemek için kilit (lock) oluşturuyoruz
        self.lock = threading.Lock()

        self.conn = sqlite3.connect(DB_PATH, check_same_thread=False)

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

            # Hızlı Arama (FTS5) Tablosu (Gelecek arama optimizasyonu için)
            self.conn.execute("""
            CREATE VIRTUAL TABLE IF NOT EXISTS pdf_search_index USING fts5(
                pdf_path,
                page,
                content
            )
            """)
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
            self.conn.execute("DELETE FROM groups WHERE id=?", (group_id,))
            self.conn.commit()

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
