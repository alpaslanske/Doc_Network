import json
import os
import sqlite3
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

import pymupdf
import markdown
from PySide6.QtCore import QCoreApplication, QPoint, QUrl

from database import Database, get_user_data_dir, migrate_legacy_database
import lang
from widgets.doc_workers import GlobalLibrarySearchWorker
from widgets.doc_viewer import PageRenderWorker, screen_to_pdf_point
from widgets.md_viewer import is_internal_anchor_url, sanitize_markdown_html
from widgets.office_cache import _content_digest, get_office_cache_path
from widgets.data_backup import create_data_backup, restore_data_backup


class OfficeCacheTests(unittest.TestCase):
    def test_translation_files_are_found_relative_to_the_resource_directory(self):
        self.assertTrue((lang.LANGS_DIR / "tr.json").is_file())

    def test_cache_key_distinguishes_same_named_sources_and_content(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            first = root / "first" / "report.docx"
            second = root / "second" / "report.docx"
            first.parent.mkdir()
            second.parent.mkdir()
            first.write_bytes(b"first")
            second.write_bytes(b"other")

            first_cache = get_office_cache_path(first, root / "cache")
            cache_hits = _content_digest.cache_info().hits
            self.assertEqual(
                get_office_cache_path(first, root / "cache"),
                first_cache,
            )
            self.assertEqual(_content_digest.cache_info().hits, cache_hits + 1)
            second_cache = get_office_cache_path(second, root / "cache")
            self.assertNotEqual(first_cache, second_cache)

            first.write_bytes(b"newer")
            stat = first.stat()
            os.utime(first, ns=(stat.st_atime_ns, stat.st_mtime_ns + 1_000_000))
            updated_cache = get_office_cache_path(first, root / "cache")
            self.assertNotEqual(first_cache, updated_cache)


class DatabaseGroupTests(unittest.TestCase):
    def test_user_data_directory_uses_the_platform_application_path(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            with patch(
                "database.QStandardPaths.writableLocation",
                return_value=temp_dir,
            ):
                self.assertEqual(get_user_data_dir(), Path(temp_dir))

    def test_legacy_database_migration_includes_uncheckpointed_wal_data(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source_path = root / "legacy" / "bookmarks.db"
            destination_path = root / "user-data" / "bookmarks.db"
            source_path.parent.mkdir()
            source = sqlite3.connect(source_path)
            source.execute("PRAGMA journal_mode=WAL")
            source.execute("CREATE TABLE sample(value TEXT)")
            source.execute("INSERT INTO sample VALUES('preserved')")
            source.commit()

            self.assertTrue(
                migrate_legacy_database(source_path, destination_path)
            )
            self.assertFalse(
                migrate_legacy_database(source_path, destination_path)
            )
            destination = sqlite3.connect(destination_path)
            try:
                self.assertEqual(
                    destination.execute(
                        "SELECT value FROM sample"
                    ).fetchone(),
                    ("preserved",),
                )
            finally:
                destination.close()
            source.close()

    def test_deleting_group_removes_its_bookmarks_and_reparents_children(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            db_path = Path(temp_dir) / "data" / "bookmarks.db"
            db = Database(db_path)

            self.addCleanup(db.conn.close)
            db.add_group("Parent")
            parent_id = db.get_groups()[0][0]
            db.add_group("Selected", parent_id)
            selected_id = next(g[0] for g in db.get_groups() if g[1] == "Selected")
            db.add_group("Child", selected_id)
            child_id = next(g[0] for g in db.get_groups() if g[1] == "Child")
            db.add_bookmark("Selected bookmark", "selected.pdf", 0, selected_id)
            db.add_bookmark("Child bookmark", "child.pdf", 1, child_id)

            db.delete_group(selected_id)

            groups = {group_id: parent for group_id, _, parent in db.get_groups()}
            self.assertNotIn(selected_id, groups)
            self.assertEqual(groups[child_id], parent_id)
            bookmarks = {bookmark[1]: bookmark[4] for bookmark in db.get_bookmarks()}
            self.assertNotIn("Selected bookmark", bookmarks)
            self.assertEqual(bookmarks["Child bookmark"], child_id)

    def test_backup_round_trip_restores_database_and_settings(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            data_dir = root / "data"
            data_dir.mkdir()
            database_path = data_dir / "bookmarks.db"
            settings_path = data_dir / "settings.json"
            backup_path = root / "backup.zip"

            original_db = Database(database_path)
            original_db.add_group("Backed up group")
            settings_path.write_text(
                '{"library_path": "/documents"}', encoding="utf-8"
            )
            create_data_backup(original_db.conn, settings_path, backup_path)

            original_db.add_group("Temporary group")
            settings_path.write_text('{"theme": "dark"}', encoding="utf-8")
            restore_data_backup(backup_path, data_dir, original_db.conn)

            restored_db = Database(database_path)
            self.addCleanup(restored_db.conn.close)
            self.assertEqual(
                [group[1] for group in restored_db.get_groups()],
                ["Backed up group"],
            )
            self.assertEqual(
                json.loads(settings_path.read_text(encoding="utf-8")),
                {"library_path": "/documents"},
            )

    def test_invalid_backup_does_not_replace_existing_data(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            data_dir = root / "data"
            data_dir.mkdir()
            database_path = data_dir / "bookmarks.db"
            settings_path = data_dir / "settings.json"
            backup_path = root / "invalid.zip"

            active_db = Database(database_path)
            active_db.add_group("Keep me")
            settings_path.write_text("{}", encoding="utf-8")
            with zipfile.ZipFile(backup_path, "w") as archive:
                archive.writestr("bookmarks.db", b"invalid")
                archive.writestr("settings.json", "{}")

            with self.assertRaises((ValueError, sqlite3.DatabaseError)):
                restore_data_backup(backup_path, data_dir, active_db.conn)

            db = Database(database_path)
            self.addCleanup(db.conn.close)
            self.assertEqual([group[1] for group in db.get_groups()], ["Keep me"])
            self.assertEqual(
                json.loads(settings_path.read_text(encoding="utf-8")), {}
            )


class GlobalLibrarySearchTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QCoreApplication.instance() or QCoreApplication([])

    def test_worker_searches_notes_and_pdf_and_reports_unreadable_files(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            pdf_path = Path(temp_dir) / "notes.pdf"
            broken_pdf = Path(temp_dir) / "broken.pdf"
            with pymupdf.open() as doc:
                page = doc.new_page()
                page.insert_text((72, 72), "needle in document")
                doc.save(pdf_path)
            broken_pdf.write_text("not a PDF", encoding="utf-8")

            bookmarks = [
                (7, "Bookmark", str(pdf_path), 0, 3, 0.0, 0.0, "needle in note")
            ]
            worker = GlobalLibrarySearchWorker(
                12,
                "needle",
                bookmarks,
                {3: "Group"},
                {str(broken_pdf)},
            )
            emissions = []
            worker.results_ready.connect(lambda *args: emissions.append(args))
            worker.run()

            self.assertEqual(len(emissions), 1)
            request_id, matches, errors = emissions[0]
            self.assertEqual(request_id, 12)
            self.assertEqual(
                {match[0][0] for match in matches},
                {"global_bookmark", "global_pdf"},
            )
            self.assertEqual(len(errors), 1)
            self.assertIn("broken.pdf", errors[0])

    def test_worker_searches_pdf_and_markdown_recursively_under_library_root(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            library_root = Path(temp_dir) / "library"
            nested_dir = library_root / "nested"
            nested_dir.mkdir(parents=True)
            pdf_path = nested_dir / "manual.pdf"
            markdown_path = library_root / "notes.md"
            with pymupdf.open() as doc:
                page = doc.new_page()
                page.insert_text((72, 72), "library needle")
                doc.save(pdf_path)
            markdown_path.write_text("library needle", encoding="utf-8")

            worker = GlobalLibrarySearchWorker(
                24, "needle", [], {}, set(), library_root
            )
            emissions = []
            worker.results_ready.connect(lambda *args: emissions.append(args))
            worker.run()

            self.assertEqual(len(emissions), 1)
            _, matches, errors = emissions[0]
            self.assertFalse(errors)
            self.assertEqual(
                {match[0][0] for match in matches},
                {"global_pdf", "global_md"},
            )


class DocumentCoordinateTests(unittest.TestCase):
    def test_screen_coordinates_round_trip_for_cropped_rotated_pdf(self):
        with pymupdf.open() as doc:
            page = doc.new_page(width=600, height=800)
            page.set_cropbox(pymupdf.Rect(30, 40, 550, 750))
            page.set_rotation(90)
            zoom = 2.0
            original = pymupdf.Point(120, 180)
            matrix = page.rotation_matrix * pymupdf.Matrix(zoom, zoom)
            crop_rect = page.cropbox * matrix
            rendered = original * matrix - crop_rect.top_left
            screen_point = QPoint(round(rendered.x), round(rendered.y))

            restored = screen_to_pdf_point(screen_point, page, zoom)

            self.assertAlmostEqual(restored.x, original.x, delta=0.5)
            self.assertAlmostEqual(restored.y, original.y, delta=0.5)

    def test_page_render_worker_returns_owned_image_data(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            pdf_path = Path(temp_dir) / "render.pdf"
            with pymupdf.open() as doc:
                doc.new_page(width=200, height=300)
                doc.save(pdf_path)

            worker = PageRenderWorker(5, str(pdf_path), 0, 1.0)
            rendered = []
            worker.rendered.connect(lambda *args: rendered.append(args))
            worker.run()

            self.assertEqual(len(rendered), 1)
            request_id, page_no, zoom, image = rendered[0]
            self.assertEqual((request_id, page_no, zoom), (5, 0, 1.0))
            self.assertEqual((image.width(), image.height()), (200, 300))


class MarkdownSafetyTests(unittest.TestCase):
    def test_raw_script_markup_is_removed_but_markdown_headings_remain(self):
        rendered = markdown.markdown(
            '<script>alert("unsafe")</script>\n\n# Safe heading',
            extensions=["toc"],
        )
        cleaned = sanitize_markdown_html(
            rendered
        )

        self.assertNotIn("<script", cleaned.lower())
        self.assertIn("<h1", cleaned)
        self.assertIn("Safe heading", cleaned)
        self.assertNotIn("onerror", sanitize_markdown_html(
            '<img src="x" onerror="alert(1)">'
        ))

    def test_fragment_only_and_same_file_links_are_internal_anchors(self):
        self.assertTrue(
            is_internal_anchor_url(QUrl("#section"), "/tmp/docs/readme.md")
        )
        self.assertTrue(
            is_internal_anchor_url(
                QUrl("file:///tmp/docs/readme.md#section"),
                "/tmp/docs/readme.md",
            )
        )
        self.assertTrue(
            is_internal_anchor_url(
                QUrl("file:///tmp/docs/#section"),
                "/tmp/docs/readme.md",
            )
        )
        self.assertFalse(
            is_internal_anchor_url(
                QUrl("https://example.com/#section"),
                "/tmp/docs/readme.md",
            )
        )


if __name__ == "__main__":
    unittest.main()
