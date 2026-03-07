# Prisma — Document Processing API

A Flask web application that converts uploaded DOCX files or rich-text editor input into professionally branded Word documents, with optional Table of Contents and Cover Page.

Built by [Codes Technology](https://www.codestechnology.com/)

---

## What It Does

1. Accepts input via **file upload** (`.docx`) or a **rich-text editor**
2. Classifies and restructures content into headings, body, and bullet points
3. Applies company branding — fonts, colors, spacing, header/footer from template
4. Optionally prepends a **Cover Page** (AI-generated summary via Ollama) and/or a **Table of Contents**
5. Outputs a downloadable `.docx` or `.pdf`

---

## Project Structure

```
Documentation_api/
├── app.py                        # Flask entry point, routes
├── document_processor.py         # Core pipeline: Y-Split Extract → Regex Engine → Brand
├── config.py                     # Folder paths, secret key, file size limit
│
├── utils/
│   ├── batch_processor.py        # Applies cover/TOC, saves final file
│   ├── style_manager.py          # Branding: fonts, sizes, colors, header/footer
│   ├── toc_manager.py            # Word-native Table of Contents field
│   ├── cover_page_manager.py     # Cover page with AI summary (Ollama)
│   └── db_manager.py             # MySQL job tracking
│
├── templates/
│   ├── upload.html               # Main UI (text editor + file upload toggle)
│   └── result.html               # Download page
│
├── static/
│   ├── css/style.css
│   └── img/
│
├── output/                       # Generated documents saved here
├── uploads/                      # Temporary upload staging
└── letter_head_1.docx            # Branding template (header/footer source)
```

---

## Setup

### Requirements

```
Flask
python-docx
mysql-connector-python
python-dotenv
ollama          # optional — for AI cover page summaries
docx2pdf        # optional — Windows only, for PDF export
```

Install:
```bash
pip install -r requirements.txt
```

### Environment Variables

Create a `.env` file in the project root:

```env
DB_USER=your_db_user
DB_PASSWORD=your_db_password
DB_HOST=localhost
DB_DATABASE=prisma_db
SECRET_KEY=your_secret_key
```

### Database

Create the jobs tracking table in MySQL:

```sql
CREATE TABLE processing_jobs (
    job_id          VARCHAR(36) PRIMARY KEY,
    original_files  TEXT,
    file_count      INT,
    status          VARCHAR(20),
    output_filename VARCHAR(255),
    error_message   TEXT,
    created_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    completed_at    TIMESTAMP NULL
);
```

### Branding Template

Place your company letterhead DOCX at the path defined by `TEMPLATE_DOCX` in `config.py`. The header and footer from this file are copied to every generated document.

---

## Running

```bash
python app.py
```

Runs on `http://0.0.0.0:5000` in debug mode.

---

## API

### `POST /api/v1/process`

Accepts `multipart/form-data`:

| Field | Type | Description |
|---|---|---|
| `text_input` | string | HTML from the rich-text editor |
| `files[]` | file | `.docx` file upload |
| `include_cover` | `"true"` / `"false"` | Prepend AI cover page |
| `include_toc` | `"true"` / `"false"` | Prepend Table of Contents |

Either `text_input` or `files[]` must be provided. Both can be provided together.

**Success response:**
```json
{
  "success": true,
  "data": {
    "filename": "Document_20260306_123456.docx",
    "redirect_url": "/result/Document_20260306_123456.docx"
  }
}
```

### `GET /download/<filename>`

Downloads the processed `.docx` file.

### `GET /download-pdf/<filename>`

Converts and downloads as `.pdf`. Requires Windows + Microsoft Word installed.

---

## Processing Pipeline

The engine abandons brittle, hardcoded English keywords in favor of a **"Y-Split" Architecture** and a **Dynamic Regex Engine**. It relies purely on structural math and pattern recognition to format documents.

### 1. Extraction (Y-Split Architecture)
* **PDFs (`adobe_helper.py`):** Sent to the **Adobe PDF Services API** to safely extract structural JSON, bypassing manual token handling.
* **DOCX (Smart XML Extractor):** Uses an `xpath` crawler to bypass Microsoft Word's hidden `<mc:Fallback>` duplicate layers. It pierces through invisible layout tables and text boxes to extract raw text perfectly, without squishing soft line breaks (`<w:br>`).

### 2. The Stitcher
Automatically detects and heals shattered numbered lists natively found in messy DOCX files (e.g., intercepting a separated `1` and `. Item` and gluing them back together into `1. Item`).

### 3. Classification (Dynamic Regex Engine)
Every extracted line is passed through math-based rules:
* **Headings (`H3`):** Any short phrase (≤ 6 words) ending in a colon (e.g., `Deliverables:`, `Tasks:`).
* **Labels:** Extracts and formats inline labels automatically (e.g., `FR-001: QR Code Scanning`).
* **Numbered Lists:** Matches regex shapes (`^\d+[\.\)]\s+[A-Z]`), allowing infinite scaling (1 to 10,000) without breaking.
* **Bullets:** Detects explicit PDF/Word characters (`•`, `-`, `☐`) or implicitly inherits list context from the section label above it.

### 4. Branding
The `StyleManager` takes the structurally classified document and applies the exact company fonts (Calibri), sizes, colors, and complex multi-level list indents from the master `TEMPLATE.docx`.

### Classification Signals (A–G)

| Signal | Rule |
|---|---|
| A | Named style is already a Heading or ListBullet — trust it |
| B | Font size ratio ≥ 1.4× body → heading1, ≥ 1.2× → heading2, ≥ 1.1× → heading3 |
| C | All-caps short line → heading1 |
| D | Bold + short line (≤ 10 words) → heading3 |
| E | Structural pattern `Word 1.2:` → heading3 |
| F | Short line (≤ 8 words) ending with `:` → heading3 |
| G | Looks like a bullet (`•`, `-`, numbered) → bullet |

---

## Key Design Decisions

**Page break instead of section break in TOC**
A section break caused all TOC page numbers to show as `1` because Word restarted the page count for each section. A simple page break keeps one section so page numbers are continuous.

**`\o "1-3"` only — no `\u` switch in TOC field**
The `\u` switch requires `outlineLvl` set on each paragraph's own `pPr`. Heading styles don't set this per-paragraph, so `\u` finds nothing. `\o "1-3"` works from style assignment alone.

**Heading character styles in branding config**
Word renders heading text through linked character styles (`Heading1Char`, etc.) which override the paragraph style. Without also setting these, headings render at the wrong size and in blue.

**`dirty=true` on TOC field**
Forces Word to recalculate page numbers every time the document is opened, without the user needing to manually update fields.

---

## Optional Features

### Cover Page (requires Ollama)

Install and run [Ollama](https://ollama.ai/) locally with the `phi3` model:

```bash
ollama pull phi3
ollama serve
```

If Ollama is not running, the cover page still generates but uses a fallback summary: *"Technical documentation and strategic project overview."*

### PDF Export (Windows only)

Requires Microsoft Word installed. Uses `docx2pdf` which calls Word via COM automation (`pythoncom`). Will return `501` on non-Windows systems.
