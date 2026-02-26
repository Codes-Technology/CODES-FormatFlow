# Document Processing & Standardization System

## Project Overview
This project is a Flask-based web application designed to automate the standardization and formatting of technical documentation. It accepts various input formats (PDF, DOCX, TXT), extracts content, attempts to reconstruct lost structure (like tables and lists), and applies a strict corporate branding template (Letterpad).

The core philosophy of the system is **"Universal Extraction -> Template Application"**. All inputs are first normalized to a raw format, then rebuilt from the ground up using a master style template.

## Architecture

### 1. The Core Pipeline (`document_processor.py`)
The `DocumentProcessor` class orchestrates the entire flow:
- **Input Normalization**: Converts PDF/TXT/DOCX into a standardized "Intermediary DOCX".
- **Structure Recovery**: Detects implicit structures (text tables, bullet lists) in raw text.
- **Style Application**: Applies the corporate styles via `StyleManager`.

### 2. Batch Processing (`utils/batch_processor.py`)
Handles multi-file workflows:
- Processes multiple documents in parallel.
- **Merging**: Combines multiple processed documents into a single Master Document.
- **Safe Merge (`utils/merge_fix.py`)**: A custom merge engine that handles lower-level XML operations to prevent corruption (specifically fixing `NumberingPart` issues common in `python-docx`).
- **PDF Conversion**: Converts final DOCX files to PDF using Microsoft Word COM automation for perfect fidelity.

### 3. Styling Engine (`utils/style_manager.py`)
The heart of the formatting logic. It is not just a "style applier" but a "document reconstructor".
- **Table Reconstruction**: Scans for text patterns that look like tables (e.g., text separated by tabs or wide gaps, "Phase - Duration" lists) and converts them into real Word tables.
- **Smart Heading Detection**: Analyzes font size, casing, and numbering patterns (e.g., "1.1 Overview", "step 1") to apply true Heading styles.
- **Branding**: Copies headers/footers from a template file (`Letter_pad.docx`) to the target document.
- **Font Normalization**: Enforces Calibri and specific font sizes (16pt for titles, 14pt for subtitles, 11pt for body).

### 4. Cover Page Generation (`utils/cover_page_manager.py`)
Automatically generates a professional cover page:
- **Content Extraction**: Pulls title and subtitle candidates from the document content.
- **AI Summary**: Uses a local Ollama model (Phi-3) to generate a 1-2 sentence executive summary of the document.
- **Vertical Centering**: Uses advanced XML manipulation (`vAlign="center"` in `sectPr`) to ensure the cover page is vertically centered regardless of content length.

---

## File Manifest & Responsibilities

### Root Directory
| File | Description |
| :--- | :--- |
| **`app.py`** | **Entry Point.** The Flask web server. Handles file uploads (`/`, `/batch`), routing, and response rendering. Manages temporary file lifecycles. |
| **`document_processor.py`** | **Controller.** Contains the `DocumentProcessor` class. Connects the extraction tools (pdf2docx) with the styling tools (StyleManager). |
| **`config.py`** | **Configuration.** Global settings: Upload folders, Allowed extensions (pdf, docx, txt), Secret keys, and Template paths. |
| **`extract_images_v2.py`** | **Utility.** A standalone script for extracting images from DOCX files (helper tool, not used in main pipeline). |

### `utils/` Directory
| File | Description |
| :--- | :--- |
| **`batch_processor.py`** | **Workflow Manager.** Handles the logic for the "Batch Merge" feature. Iterates through files, processes them, and calls `merge_fix` to combine them. |
| **`style_manager.py`** | **The "Brain".** Contains 1400+ lines of formatting logic. Handles: Table reconstruction, List detection, Font fixing, Header/Footer copying, and Section property management. |
| **`cover_page_manager.py`** | **Feature.** Generates the cover page. Features a robust specific XML fix to ensure "True Vertical Centering" by manipulating the first section break's properties. |
| **`merge_fix.py`** | **Stability Fix.** Contains `safe_merge()` and `inject_numbering_part()`. Solves critical bugs in the `docxcompose` library by manually injecting missing XML parts required for merging lists. |
| **`pdf_converter.py`** | **EXPORTER.** Uses `win32com` to control an installed instance of Microsoft Word to save DOCX as PDF. This ensures the PDF looks *exactly* like the Word doc (unlike python libraries which often mess up formatting). |
| **`table_detector.py`** | **Helper.** Uses `pdfplumber` to detect if a PDF contains tables, guiding the extraction engine to use the right strategy. |
| **`file_validator.py`** | **Security.** Validates file extensions and sanitizes filenames to prevent path traversal attacks. |

---

## Detailed Logic Flows

### The "Safe Merge" Strategy (`merge_fix.py`)
Merging Word documents is notoriously difficult because each document has its own `document.xml` and `numbering.xml` (list definitions).
1. **Pre-processing**: We define a "Master" document (the template).
2. **Injection**: We assume incoming docs might be broken/missing XML parts. `inject_numbering_part()` forces a valid numbering schema into them.
3. **Composition**: We use `docxcompose` to merge.
4. **Fallback**: If `docxcompose` fails, we fall back to a "Manual Append" mode where we copy paragraphs one-by-one, skipping complex section properties to avoid crashing.

### The "Vertical Centering" specific (`cover_page_manager.py`)
Word does not have a simple "Vertical Center" button for a single page. It is a Section Property.
- We create the cover page content.
- We insert a **Section Break (Next Page)** at the end of the cover content.
- **CRITICAL**: We apply `<w:vAlign w:val="center"/>` to the **Section Properties (`sectPr`)** of that break. This tells Word "The section *ending* here should be vertically centered".
- This logic was isolated into `_add_cover_section_break()` to prevent StyleManager's general formatting (which forces top-alignment) from overwriting it.

### The "Call Order" Importance (`style_manager.py`)
The order of operations in `apply_template_styles` is critical:
1. Styles are standardized first.
2. Tables and Lists are reconstructed (converting raw text → structures).
3. Fonts and Colors are normalized.
4. **LAST STEP**: `create_cover_page()` is called. This is done *last* so that no subsequent global formatting sweep accidentally resets the unique alignment properties of the cover page.
