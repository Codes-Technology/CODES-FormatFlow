"""
DocumentProcessor — Signal-Based Classifier + Fresh Build

PIPELINE (file upload path):
  1. EXTRACT  — read signals from ORIGINAL file XML before anything changes
  2. CLASSIFY — size ratio + named style + numId + bold (no hardcoded keywords)
  3. POST-PROCESS — context inheritance + split crammed list items
  4. BUILD    — fresh doc from template with correct styles
  5. BRAND    — StyleManager applies color, font, sizes, header/footer

PIPELINE (editor/text input path):
  html_to_docx() — HTML already has structure, skip classify pipeline entirely
"""

import os
import re
import traceback
from statistics import mode, StatisticsError
from docx import Document
from pdf2docx import Converter
from docx.document import Document as _Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from copy import deepcopy
from bs4 import BeautifulSoup
from utils.style_manager import StyleManager
from config import TEMPLATE_DOCX

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

# Bullet characters that pdf2docx leaves as plain text after PDF conversion
BULLET_PREFIXES = ('•', '▪', '►', '➤', '→', '○', '●', '‣', '□', '■', '◆', '✓', '✗', '–', '—', '☐', '☑', '☒')


class DocumentProcessor:
    def __init__(self, template_path: str = TEMPLATE_DOCX):
        self.template_path = template_path
        self.style_manager = StyleManager(template_path)

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # PUBLIC API — called by app.py
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    def universal_extract(self, input_path: str, output_path: str) -> dict:
        """Entry point for uploaded files. Routes by extension."""
        try:
            ext = os.path.splitext(input_path)[1].lower()
            if ext == '.pdf':
                return self._pipeline_pdf(input_path, output_path)
            elif ext == '.docx':
                return self._pipeline_docx(input_path, output_path)
            return {'success': False, 'error': f'Unsupported type: {ext}'}
        except Exception as e:
            traceback.print_exc()
            return {'success': False, 'error': str(e)}

    def html_to_docx(self, html: str) -> Document:
        """
        Rich-text editor HTML → branded DOCX.
        HTML already encodes structure (h1/h2/ul/p) so we skip the classifier.
        """
        doc = Document(self.template_path)
        soup = BeautifulSoup(html, 'html.parser')
        for el in soup.find_all(['h1', 'h2', 'h3', 'h4', 'p', 'ul', 'ol', 'table']):
            text = el.get_text().strip()
            if el.name in ('h1', 'h2', 'h3', 'h4'):
                if text:
                    doc.add_heading(text, level=int(el.name[1]))
            elif el.name == 'p':
                if text:
                    doc.add_paragraph(text)
            elif el.name in ('ul', 'ol'):
                style = 'List Bullet' if el.name == 'ul' else 'List Number'
                for li in el.find_all('li', recursive=False):
                    li_text = li.get_text().strip()
                    if li_text:
                        doc.add_paragraph(li_text, style=style)
            elif el.name == 'table':
                self._add_html_table(doc, el)
        return self.style_manager.apply_template_styles(doc)

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # DOCUMENT ITERATORS
    # Used by _extract_blocks to walk all paragraphs including tables
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    def iter_block_items(self, parent):
        """Yield paragraphs and tables in document order."""
        parent_elm = parent.element.body if isinstance(parent, _Document) else parent._element
        for child in parent_elm.iterchildren():
            if child.tag.endswith('}p'):
                yield Paragraph(child, parent)
            elif child.tag.endswith('}tbl'):
                yield Table(child, parent)

    def iter_all_blocks(self, doc):
        """Iterate through paragraphs and tables, including inside table cells."""
        for block in self.iter_block_items(doc):
            yield block
            if isinstance(block, Table):
                for row in block.rows:
                    for cell in row.cells:
                        for inner in self.iter_block_items(cell):
                            yield inner

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # STEP 1 — EXTRACT
    # Read ALL signals from the ORIGINAL file before anything changes.
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    def _extract_blocks(self, docx_path: str) -> list:
        doc = Document(docx_path)

        # Detect body font size = mode of all explicit run sizes in document.
        # This becomes the baseline for ratio-based heading detection in Step 2.
        all_sizes = []
        for block in self.iter_all_blocks(doc):
            if isinstance(block, Paragraph):
                for r in block._p.findall(qn('w:r')):
                    rPr = r.find(qn('w:rPr'))
                    if rPr is not None:
                        sz = rPr.find(qn('w:sz'))
                        if sz is not None:
                            try:
                                all_sizes.append(int(sz.get(qn('w:val'))))
                            except Exception:
                                pass

        reasonable_sizes = [s for s in all_sizes if s >= 20]
        try:
            body_size = mode(reasonable_sizes) if reasonable_sizes else 22
        except StatisticsError:
            body_size = 22
        print(f'  [Classifier] Body size: {body_size} half-pts = {body_size / 2}pt')

        raw_blocks = []

        for element in doc.element.body:
            tag = element.tag.split('}')[-1] if '}' in element.tag else element.tag

            # Tables are copied raw — no classification needed
            if tag == 'tbl':
                raw_blocks.append({
                    'type': 'table', 'text': '', 'runs': [],
                    'level': 0, 'table_element': deepcopy(element)
                })
                continue
            if tag != 'p':
                continue

            text = ''.join(t.text or '' for t in element.iter(f'{{{W}}}t')).strip()
            if not text:
                continue

            # ── Paragraph-level signals ──
            pPr = element.find(qn('w:pPr'))
            named_style = num_id = None
            indent_level = 0

            if pPr is not None:
                pStyle = pPr.find(qn('w:pStyle'))
                if pStyle is not None:
                    named_style = pStyle.get(qn('w:val'), '')
                numPr = pPr.find(qn('w:numPr'))
                if numPr is not None:
                    n  = numPr.find(qn('w:numId'))
                    il = numPr.find(qn('w:ilvl'))
                    try: num_id = int(n.get(qn('w:val'))) if n is not None else None
                    except Exception: pass
                    try: indent_level = int(il.get(qn('w:val'))) if il is not None else 0
                    except Exception: pass

            # ── Run-level signals (size + bold for classifier) ──
            run_sizes, is_bold = [], False
            for r in element.findall(qn('w:r')):
                rPr = r.find(qn('w:rPr'))
                if rPr is not None:
                    sz = rPr.find(qn('w:sz'))
                    if sz is not None:
                        try: run_sizes.append(int(sz.get(qn('w:val'))))
                        except Exception: pass
                    if rPr.find(qn('w:b')) is not None:
                        is_bold = True

            run_size = max(run_sizes) if run_sizes else None

            # ── Inline runs (bold/italic preserved into output) ──
            runs = []
            for r in element.findall(qn('w:r')):
                rt = ''.join(t.text or '' for t in r.findall(qn('w:t')))
                if not rt:
                    continue
                rPr = r.find(qn('w:rPr'))
                bold = italic = False
                if rPr is not None:
                    bold   = rPr.find(qn('w:b')) is not None
                    italic = rPr.find(qn('w:i')) is not None
                runs.append({'text': rt, 'bold': bold, 'italic': italic})
            
            #--bullet prefix-stripping
            stripped_text = text
            is_pdf_bullet = False
            if text.startswith(BULLET_PREFIXES):
                stripped_text = text[1:].lstrip()
                is_pdf_bullet = True
            elif re.match(r'^[-\*]\s+\S', text):
                stripped_text = text[1:].lstrip()
                is_pdf_bullet = True

            if is_pdf_bullet:
                # Also strip the prefix from runs
                clean_runs = []
                prefix_remaining = len(text) - len(stripped_text)
                consumed = 0
                for rd in runs:
                    if consumed < prefix_remaining:
                        skip = min(prefix_remaining - consumed, len(rd['text']))
                        remaining = rd['text'][skip:].lstrip() if consumed + skip >= prefix_remaining else rd['text'][skip:]
                        consumed += skip
                        if remaining:
                            clean_runs.append({'text': remaining, 'bold': rd['bold'], 'italic': rd['italic']})
                    else:
                        clean_runs.append(rd)
                btype = 'bullet'
                raw_blocks.append({'type': btype, 'text': stripped_text, 'runs': clean_runs, 'level': indent_level})
            else:
                btype = self._classify(named_style, num_id, run_size, body_size, is_bold, text)
                raw_blocks.append({'type': btype, 'text': text, 'runs': runs, 'level': indent_level})

        return self._post_process(raw_blocks)

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # STEP 2 — CLASSIFY
    # Signal priority A→G. No hardcoded domain keywords.
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    def _classify(self, named_style, num_id, run_size, body_size, is_bold, text) -> str:
        # A — Named Word style (most reliable — trust the original author's intent)
        if named_style:
            ns = named_style.lower().replace(' ', '')
            if   'heading1' in ns: return 'heading1'
            elif 'heading2' in ns: return 'heading2'
            elif 'heading3' in ns: return 'heading3'
            elif 'heading4' in ns: return 'heading4'
            elif 'title'    in ns: return 'title'
            elif 'listparagraph' in ns or 'listbullet' in ns:
                # "Deliverables:" styled as list but it's actually a section label
                if text.rstrip().endswith(':') and len(text.split()) <= 8:
                    return 'heading3'
                return 'bullet'
            elif 'listnumber' in ns:
                return 'numbered'

        # B — Explicit list marker in XML
        if num_id is not None and num_id > 0:
            return 'bullet'

        # C — Font size ratio (core classifier — works for ANY document)
        # Uses ratio not absolute size so it adapts to any base font
        if run_size is not None and body_size > 0:
            ratio = run_size / body_size
            if   ratio >= 1.5:  return 'title'
            elif ratio >= 1.15: return 'heading2'
            elif ratio >= 1.05: return 'heading3'

        # D — Bold short line with no sentence-ending period
        if is_bold and len(text.split()) <= 12 and not text.endswith('.'):
            return 'heading3'

        # E — ALL CAPS short line
        if text.isupper() and 2 < len(text) and len(text.split()) <= 15:
            return 'heading2'

        # F — Line ending with ':' (section label) or starting with digit (numbered section)
        stripped = text.rstrip()
        if stripped.endswith(':') and len(stripped.split()) <= 8:
            return 'heading3'
        if re.match(r'^\d+[\.\)]\s+\w', text) and len(text.split()) <= 6:
            return 'heading3'

        # G — "Word N: description" e.g. "Week 1:", "Module 3:", "Unit 7.2:"
        # Structural pattern — no hardcoded words, matches any TitleCase + digit + colon
        if re.match(r'^[A-Z][a-z]+\s+[\d][\d\-\.]*\s*:', text) and len(text.split()) <= 12:
            return 'heading2'

        return 'body'

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # STEP 3 — POST-PROCESS
    # Fix misclassifications using context between adjacent blocks.
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    def _post_process(self, raw_blocks: list) -> list:
        blocks = []
        last_h3_colon   = False   # True when previous block was heading3 ending ':'
        last_was_bullet = False

        for block in raw_blocks:
            btype = block['type']
            text  = block['text']

            if btype == 'table':
                blocks.append(block)
                last_h3_colon = last_was_bullet = False
                continue

            # ── "Cloud kitchen Observe workflows:" ──
            # heading3 crammed with a trailing bullet item before the label.
            # Split into: bullet("Cloud kitchen") + heading3("Observe workflows:")
            if btype == 'heading3' and last_was_bullet:
                parts = self._try_split_crammed(text)
                if len(parts) >= 2 and parts[-1].rstrip().endswith(':'):
                    for part in parts[:-1]:
                        nb = dict(block)
                        nb['type'] = 'bullet'
                        nb['text'] = part
                        nb['runs'] = [{'text': part, 'bold': False, 'italic': False}]
                        blocks.append(nb)
                    nb = dict(block)
                    nb['text'] = parts[-1]
                    nb['runs'] = [{'text': parts[-1], 'bold': True, 'italic': False}]
                    blocks.append(nb)
                    last_h3_colon   = True
                    last_was_bullet = False
                    continue

            # ── Body paragraph after heading3+':' or after bullets → promote to bullet ──
            if btype == 'body' and (last_h3_colon or last_was_bullet):

                # Try crammed split first (e.g. "Item1 Item2 Item3" → 3 bullets)
                split_items = self._try_split_crammed(text) if last_h3_colon else []
                if not split_items and self._looks_like_list_item(text):
                    split_items = self._try_split_crammed(text)

                if split_items and len(split_items) > 1:
                    for item in split_items:
                        new_block = dict(block)
                        new_block['type'] = 'bullet'
                        new_block['text'] = item
                        new_block['runs'] = [{'text': item, 'bold': False, 'italic': False}]
                        blocks.append(new_block)
                    last_was_bullet = True
                    last_h3_colon   = False
                    continue

                # Continue bullet list if line looks like a list item
                if last_was_bullet:
                    clean = text.strip()
                    if clean and not clean.endswith('.') and not clean.endswith(':') and clean[0].isupper():
                        block['type'] = 'bullet'

            btype = block['type']   # may have changed above
            blocks.append(block)

            # Update context flags for next iteration
            if btype in ('heading1', 'heading2', 'heading3', 'title'):
                last_h3_colon   = (btype == 'heading3' and text.rstrip().endswith(':'))
                last_was_bullet = False
            elif btype in ('bullet', 'numbered'):
                last_h3_colon   = False
                last_was_bullet = True
            else:
                last_h3_colon   = False
                last_was_bullet = False

        return blocks

    def _looks_like_list_item(self, text: str) -> bool:
        """Short line without sentence structure → probably a list item."""
        if len(text) > 400:
            return False
        if re.search(r'\.\s+[A-Z]', text):   # multiple sentences → body paragraph
            return False
        if not text.endswith('.'):
            return False
        return False

    def _try_split_crammed(self, text: str) -> list:
      
        parts = re.split(r'(?<=[a-z\)\d])\s+(?=[A-Z][a-z])', text)
        if len(parts) >= 2 and all(len(p.strip()) > 2 for p in parts):
            return [p.strip() for p in parts if p.strip()]
        return []


    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # STEP 4 — BUILD
    # Write classified blocks into a fresh doc opened from template.
    # Template gives us header/footer/page layout for free.
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    def _build_fresh_doc(self, blocks: list) -> Document:
        doc = Document(self.template_path)

        for block in blocks:
            btype = block['type']
            text  = block['text']
            runs  = block['runs']

            if btype == 'table':
                doc.element.body.append(block['table_element'])
                doc.element.body.append(OxmlElement('w:p'))   # empty separator after table
                continue

            if btype in ('title', 'heading1'):
                p = doc.add_heading(text, level=1)
            elif btype == 'heading2':
                p = doc.add_heading(text, level=2)
            elif btype == 'heading3':
                p = doc.add_heading(text, level=3)
            elif btype in ('bullet', 'numbered'):
                style = 'List Bullet' if btype == 'bullet' else 'List Number'
                try:
                    p = doc.add_paragraph(style=style)
                except Exception:
                    p = doc.add_paragraph()
                self._write_runs(p, runs, text, force_no_bold=True)
            else:
                p = doc.add_paragraph()
                self._write_runs(p, runs, text, force_no_bold=True)

        return doc

    def _write_runs(self, para, runs: list, fallback_text: str, force_no_bold: bool = False):
        """Write runs preserving bold/italic. Do NOT strip text."""

        if runs:
            for rd in runs:
                text = rd['text']

                if text and text.strip():
                    r = para.add_run(text)
                    r.bold   = False if force_no_bold else rd.get('bold', False)
                    r.italic = False if force_no_bold else rd.get('italic', False)
        else:
            para.add_run(fallback_text)

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # PIPELINES
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    def _pipeline_docx(self, path: str, out: str) -> dict:
        print(f'  [DocProcessor] {os.path.basename(path)}')

        blocks = self._extract_blocks(path)              # Steps 1 + 2 + 3

        h  = sum(1 for b in blocks if b['type'].startswith('heading'))
        bl = sum(1 for b in blocks if b['type'] == 'bullet')
        t  = sum(1 for b in blocks if b['type'] == 'table')
        print(f'  [Classifier] {len(blocks)} blocks → {h} headings, {bl} bullets, {t} tables')

        fresh   = self._build_fresh_doc(blocks)                     # Step 4
        branded = self.style_manager.apply_template_styles(fresh)   # Step 5
        branded.save(out)

        return {'success': True, 'paragraphs': len(branded.paragraphs),
                'tables': t, 'tables_found': t}

    def _pipeline_pdf(self, path: str, out: str) -> dict:
        """PDF → pdf2docx → your perfect classifier pipeline"""
        temp = out.replace('.docx', '_raw.docx')
        
        print(f'  [DocProcessor] PDF → temp DOCX via pdf2docx: {os.path.basename(temp)}')
        
        try:
            
            cv = Converter(path)
            cv.convert(temp, multi_processing=True, OCR_language="eng")
            cv.close()
        except Exception as e:
            return {'success': False, 'error': f'pdf2docx failed: {e}'}
        
        # Your classifier magic happens here
        return self._pipeline_docx(temp, out)

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # HTML TABLE HELPER
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    def _add_html_table(self, doc: Document, element):
        """Converts a BeautifulSoup <table> element into a Word table."""
        rows = element.find_all('tr')
        if not rows:
            return
        cols = max(len(r.find_all(['td', 'th'])) for r in rows)
        if not cols:
            return
        table = doc.add_table(rows=len(rows), cols=cols)
        try: table.style = 'Table Grid'
        except Exception: pass
        for i, row in enumerate(rows):
            for j, cell in enumerate(row.find_all(['td', 'th'])):
                if j < cols:
                    table.rows[i].cells[j].text = cell.get_text().strip()