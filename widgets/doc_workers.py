#!/usr/bin/env python3
import os
from pathlib import Path

from PySide6.QtCore import QThread, Signal
import pymupdf
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
from PySide6.QtGui import QImage, QPainter
from PySide6.QtPrintSupport import QPrinter


class PDFSearchWorker(QThread):
    results_ready = Signal(list)
    error = Signal(str)

    def __init__(self, file_path, search_text):
        super().__init__()
        self.file_path = file_path
        self.search_text = search_text

    def run(self):
        try:
            results = []
            # 'with' bloğu PDF'i açar ve işlem bitince (veya hata çıkarsa) otomatik kapatır
            with pymupdf.open(self.file_path) as doc:
                for page_no in range(len(doc)):
                    if self.isInterruptionRequested():
                        break
                    page = doc[page_no]
                    for rect in page.search_for(self.search_text):
                        results.append((page_no, rect))

            self.results_ready.emit(results)
        except Exception as e:
            self.error.emit(str(e))


class GlobalLibrarySearchWorker(QThread):
    results_ready = Signal(int, list, list)
    error = Signal(int, str)

    def __init__(
        self, request_id, query, bookmarks, groups, document_paths,
        library_path=None,
    ):
        super().__init__()
        self.request_id = request_id
        self.query = query
        self.bookmarks = bookmarks
        self.groups = groups
        self.document_paths = set(document_paths)
        self.library_path = Path(library_path) if library_path else None

    def run(self):
        try:
            query_lower = self.query.lower()
            matches = []
            errors = []

            if self.library_path is not None:
                if not self.library_path.is_dir():
                    errors.append(
                        f"{self.library_path}: kütüphane klasörü bulunamadı"
                    )
                else:
                    walk_errors = []
                    for current_dir, directories, filenames in os.walk(
                        self.library_path,
                        onerror=walk_errors.append,
                        followlinks=False,
                    ):
                        if self.isInterruptionRequested():
                            return

                        errors.extend(str(error) for error in walk_errors)
                        walk_errors.clear()
                        current_path = Path(current_dir)
                        directories[:] = [
                            directory
                            for directory in directories
                            if not (current_path / directory).is_symlink()
                        ]
                        for filename in filenames:
                            if self.isInterruptionRequested():
                                return
                            file_path = current_path / filename
                            if file_path.is_symlink():
                                continue
                            if file_path.suffix.lower() in (
                                ".pdf", ".md", ".markdown"
                            ):
                                self.document_paths.add(str(file_path))

            for bookmark in self.bookmarks:
                if self.isInterruptionRequested():
                    return

                bid, title, pdf_path, page, group_id = bookmark[:5]
                note = bookmark[7] or "" if len(bookmark) > 7 else ""

                if pdf_path:
                    self.document_paths.add(pdf_path)

                if query_lower in (title or "").lower() or query_lower in note.lower():
                    group_name = self.groups.get(group_id, "Genel")
                    file_name = Path(pdf_path).name if pdf_path else ""
                    preview = f"📌 [{group_name}] {title}\n    📄 {file_name} (Sayfa {page + 1})"
                    if query_lower in note.lower():
                        snippet = note.replace("\n", " ")[:60]
                        preview += f"\n    📝 Not: {snippet}..."
                    matches.append((("global_bookmark", bid), preview))

            for file_path in self.document_paths:
                if self.isInterruptionRequested():
                    return

                doc_path = Path(file_path)
                suffix = doc_path.suffix.lower()

                # Markdown dosyalarını tüm kütüphane aramasına dahil et
                if suffix in [".md", ".markdown"]:
                    if not doc_path.is_file():
                        errors.append(f"{doc_path}: dosya bulunamadı")
                        continue
                    try:
                        text = doc_path.read_text(encoding="utf-8")
                        occ_idx = 0
                        for idx, line in enumerate(text.splitlines()):
                            if self.isInterruptionRequested():
                                return
                            if query_lower in line.lower():
                                line_no = idx + 1
                                snippet = line.strip()[:55]
                                preview = (
                                    f"📝 {doc_path.name} (Satır {line_no})\n"
                                    f"    🔍 {snippet}..."
                                )
                                matches.append((
                                    ("global_md", str(doc_path), occ_idx, self.query),
                                    preview
                                ))
                                occ_idx += 1
                    except Exception as e:
                        errors.append(f"{doc_path}: {e}")
                    continue

                if suffix != ".pdf":
                    continue
                if not doc_path.is_file():
                    errors.append(f"{doc_path}: dosya bulunamadı")
                    continue

                try:
                    with pymupdf.open(str(doc_path)) as doc:
                        for page_no in range(len(doc)):
                            if self.isInterruptionRequested():
                                return
                            rects = doc[page_no].search_for(self.query)
                            if rects:
                                preview = (
                                    f"📄 {doc_path.name} (Sayfa {page_no + 1})\n"
                                    f"    🔍 {len(rects)} eşleşme bulundu"
                                )
                                matches.append((
                                    ("global_pdf", str(doc_path), page_no, rects),
                                    preview
                                ))
                except Exception as e:
                    errors.append(f"{doc_path}: {e}")


            self.results_ready.emit(self.request_id, matches, errors)
        except Exception as e:
            self.error.emit(self.request_id, str(e))


class TableExportWorker(QThread):
    exported = Signal(str)
    error = Signal(str)
    no_table = Signal()

    def __init__(self, file_path, page_no, x0, y0, x1, y1, save_filename):
        super().__init__()
        self.file_path = file_path
        self.page_no = page_no
        self.x0 = x0
        self.y0 = y0
        self.x1 = x1
        self.y1 = y1
        self.save_filename = save_filename

    def run(self):
        try:
            final_data = []

            # 1. Aşama: PDF'ten veriyi güvenlice oku ve kapat
            with pymupdf.open(self.file_path) as doc:
                page = doc[self.page_no]
                rect = pymupdf.Rect(self.x0, self.y0, self.x1, self.y1)
                tables = page.find_tables(clip=rect)

                if not tables or not tables.tables:
                    self.no_table.emit()
                    return

                data = tables.tables[0].extract()
                if not data:
                    self.no_table.emit()
                    return

                df = pd.DataFrame(data).fillna('').astype(str).replace(['NaN', 'None'], '', regex=True)
                cleaned_rows = df.apply(
                    lambda row: [str(val).replace('\n', ' ').strip() for val in row if str(val).strip() != ''],
                    axis=1
                ).tolist()

                if not cleaned_rows:
                    self.no_table.emit()
                    return

                max_cols = max(len(row) for row in cleaned_rows if row)

                for i, row in enumerate(cleaned_rows):
                    if not row:
                        final_data.append([""] * max_cols)
                        continue
                    if len(row) < max_cols:
                        if i == 0:
                            row = [""] * (max_cols - len(row)) + row
                        else:
                            row = row + [""] * (max_cols - len(row))
                    final_data.append(row)

            # 2. Aşama: Excel kaydetme işlemleri (Bu noktada PDF dosyası kilitli değil!)
            df_clean = pd.DataFrame(final_data)
            df_clean.to_excel(self.save_filename, index=False, header=False)

            try:
                wb = load_workbook(self.save_filename)
                ws = wb.active
                serif_font = Font(name="Times New Roman", size=11)
                for row in ws.iter_rows():
                    for cell in row:
                        cell.font = serif_font
                wb.save(self.save_filename)
            except Exception:
                pass

            self.exported.emit(self.save_filename)

        except Exception as e:
            self.error.emit(str(e))


class PDFPrintWorker(QThread):
    printed = Signal(str)
    error = Signal(str)

    def __init__(self, file_path, printer, from_page, to_page):
        super().__init__()
        self.file_path = file_path
        self.printer = printer
        self.from_page = from_page
        self.to_page = to_page

    def run(self):
        try:
            painter = QPainter(self.printer)

            # 'with' kullanarak PDF işlemlerini garantiliyoruz
            with pymupdf.open(self.file_path) as doc:
                for i in range(self.from_page, self.to_page + 1):
                    if i > self.from_page:
                        self.printer.newPage()

                    page = doc[i]
                    matrix = pymupdf.Matrix(3, 3)
                    pix = page.get_pixmap(matrix=matrix)
                    image = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)

                    rect = self.printer.pageRect(QPrinter.DevicePixel)
                    painter.drawImage(rect, image)

            painter.end()
            self.printed.emit(self.printer.printerName())

        except Exception as e:
            self.error.emit(str(e))
