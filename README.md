# PDF Network: Centralized Bookmarks for PDF Documents
[![License](https://img.shields.io/badge/license-GNU/GPLv3-blue)](LICENSE)
[![Python3](https://img.shields.io/badge/Python-3.x-red)](https://github.com/alpaslanske/Doc_Network)
[![Made In Turkey](https://github.com/alpaslanske/PDF_Network/blob/main/.assets/madeinturkiye.svg)](https://goturkiye.com/)
[![Docs](https://img.shields.io/badge/docs-online-brightgreen)](https://github.com/alpaslanske/Doc_Network/tree/main/docs)
# 📚 Doc Network

[🇬🇧 English](README.md) | [🇹🇷 Türkçe](README.tr.md)

**Doc Network** is a modern, cross-platform document viewer and personal knowledge base application that allows you to create **cross-document links (`[[file#page]]`)**, organize hierarchical bookmark groups, and perform full-text searches across your entire library.

---

## ✨ Key Features

* **Wide File Format Support:**
  * **PDF & E-Books:** `.pdf`, `.epub`, `.mobi`, `.fb2`, `.cbz`
  * **Office Documents (with Smart Caching):** `.docx`, `.odt`, `.xlsx`, `.ods`, `.pptx`, `.odp`, `.odg`, `.odf`, `.rtf`, `.csv`
* **Cross-Document Linking (`[[file#page]]`):**
  * Insert clickable links inside your study notes pointing to specific pages of other documents.
  * Clicking a link opens the target document in a new tab and jumps directly to the exact page.
* **Browser-Style Navigation History (`◀` / `▶`):**
  * Automatically records your reading position when jumping via cross-links, bookmarks, table of contents, or search results, allowing one-click back and forward navigation.
* **Dual-Mode Smart Search (Asynchronous):**
  * **Current Document Mode:** Searches the active document in the background without freezing the UI and highlights matches on the page.
  * **Entire Library & Notes Mode:** Searches across all bookmark titles, study notes in the database, and the full text of all PDF files in your library.
* **PDF Annotation & Data Extraction:**
  * Highlight selected regions and attach sticky page notes directly onto PDF pages.
  * Export selected table areas into **Excel (`.xlsx`)** spreadsheets with a single click.
  * Copy text from scanned documents or images using **OCR (Tesseract)** (supports Turkish, Polish, and English).
  * Export bookmarks and study notes into a structured **Markdown (`.md`)** summary.
* **Multi-Tab Workspace & Dashboard:**
  * Track recently opened documents along with their last read page and view multiple documents simultaneously in tabs.
  * Built-in English (`🇬🇧 EN`) and Turkish (`🇹🇷 TR`) interface languages.

---

## 🚀 Installation & Usage

### 1. Portable Releases (Recommended)
Download the pre-built package for your operating system from the **Releases** page and run it without installation:
* **Linux:** `Doc_Network_Portable-x86_64.AppImage` (Make executable and run: `chmod +x Doc_Network_Portable-x86_64.AppImage`)
* **Windows:** `Doc_Network_Portable.exe`

### 2. Running from Source

1. Clone the repository:
   ```bash
   git clone https://github.com
   cd Doc_Network
   ```

2. Install the required Python dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Launch the application:
   ```bash
   python3 doc_network.py
   ```

> **Note (Optional System Dependencies):**
> * To view Office documents (`.docx`, `.xlsx`, `.odt`, etc.), **LibreOffice (`soffice`)** must be installed on your system.
> * To use image-to-text OCR copying, **Tesseract OCR** (with `eng`, `tur`, `pol` language packs) must be installed on your system.

---

## 🖱️ Quick Tips

* **Right-Click Panning:** Hold the right mouse button on a document to drag and pan around the page.
* **Left-Click Area Selection:** Drag with the left mouse button to select an area, then right-click to copy text, add a bookmark, highlight, or export a table to Excel.
* **Zooming:** Hold `Ctrl + Mouse Wheel` to zoom in and out smoothly.
* **Editing Notes:** Double-click the note area in the bottom-left panel to edit your notes and add `[[filename.pdf#15]]` links.

---

## 🛠️ Tech Stack

* **GUI Framework:** Python 3, PySide6 (Qt6)
* **Document Engine:** PyMuPDF (`fitz`), LibreOffice Headless
* **Database:** SQLite3
* **OCR & Data Export:** Tesseract OCR, Pillow, Pandas, OpenPyXL
