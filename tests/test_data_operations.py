import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import pymupdf
from PySide6.QtCore import QCoreApplication

import database
import lang
from database import Database
from widgets.doc_workers import GlobalLibrarySearchWorker
from widgets.office_cache import get_office_cache_path


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
            second_cache = get_office_cache_path(second, root / "cache")
            self.assertNotEqual(first_cache, second_cache)

            first.write_bytes(b"newer")
            updated_cache = get_office_cache_path(first, root / "cache")
            self.assertNotEqual(first_cache, updated_cache)


class DatabaseGroupTests(unittest.TestCase):
    def test_deleting_group_removes_its_bookmarks_and_reparents_children(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            db_path = Path(temp_dir) / "data" / "bookmarks.db"
            with patch.object(database, "DB_PATH", db_path):
                db = Database()

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


if __name__ == "__main__":
    unittest.main()
