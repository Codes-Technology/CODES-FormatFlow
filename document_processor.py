"""
DocumentProcessor — Signal-Based Classifier

"""

import os
import re
import time
import traceback
import tempfile
from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from bs4 import BeautifulSoup
from utils.style_manager import StyleManager
from utils.adobe_helper import adobe_pdf_extract
from utils.toc_manager import TocManager
from utils.cover_page_manager import CoverPageManager
from config import TEMPLATE_DOCX, ADOBE_CLIENT_ID, ADOBE_CLIENT_SECRET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


class DocumentProcessor:
    def __init__(self, template_path: str, font_family: str = 'Calibri', font_size: int = 11, 
                 include_cover: bool = False, include_toc: bool = False):
        self.template_path = template_path
        self.style_manager = StyleManager(template_path, font_family, font_size)
        self.include_cover = include_cover
        self.include_toc = include_toc
        self.toc_manager = TocManager()
        self.cover_manager = CoverPageManager()

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # PUBLIC API
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    def universal_extract(self, input_path: str, output_path: str) -> dict:
        """Entry point for uploaded files. Routes by extension."""
        try:
            ext = os.path.splitext(input_path)[1].lower()
            if ext == '.pdf':
                return self._pipeline_pdf(input_path, output_path)
            elif ext == '.docx':
                return self._pipeline_docx(input_path, output_path)
            return {'success': False, 'error': f'Unsupported file type: {ext}'}
        except Exception as e:
            traceback.print_exc()
            return {'success': False, 'error': str(e)}

    def html_to_docx(self, html: str) -> Document:
        doc = Document(self.template_path)
        soup = BeautifulSoup(html, 'html.parser')

        print(f"[DocumentProcessor] Incoming HTML Length: {len(html)}")

        processed_elements = set()

        # ── FIX: pass doc explicitly into process_node ──
        def process_node(node, container):
            """
            Recursive HTML walker that preserves formatting.
            container: Can be a Document or a Paragraph object.
            """
            if node in processed_elements:
                return

            if not getattr(node, 'name', None):
                # Text node
                clean_text = str(node).strip()
                if clean_text:
                    if isinstance(container, Paragraph):
                        container.add_run(clean_text)
                    else:
                        container.add_paragraph(clean_text)
                return

            # --- Block Elements ---
            if node.name in ('h1', 'h2', 'h3', 'h4'):
                level = int(node.name[1])
                para = container.add_heading('', level=level)
                for child in node.contents:
                    process_node(child, para) # Pass the paragraph to continue run building
                processed_elements.add(node)

            elif node.name == 'p':
                para = container.add_paragraph()
                for child in node.contents:
                    process_node(child, para)
                processed_elements.add(node)

            elif node.name == 'table':
                self._add_html_table(container, node)
                processed_elements.add(node)
                # Tables handle their own children in _add_html_table

            elif node.name in ('ul', 'ol'):
                style = 'List Bullet' if node.name == 'ul' else 'List Number'
                for li in node.find_all('li', recursive=False):
                    para = container.add_paragraph(style=style)
                    for child in li.contents:
                        process_node(child, para)
                processed_elements.add(node)

            # --- Inline/Formatting Elements ---
            elif node.name in ('b', 'strong', 'i', 'em', 'u', 'span'):
                if isinstance(container, Paragraph):
                    # We are already inside a paragraph, add a run and recurse
                    run = container.add_run()
                    if node.name in ('b', 'strong'): run.bold = True
                    if node.name in ('i', 'em'): run.italic = True
                    if node.name == 'u': run.underline = True
                    
                    # Process children into THIS run if they are simple text, 
                    # else recurse to handle nested formatting like <b><i>...</i></b>
                    for child in node.contents:
                        if not getattr(child, 'name', None):
                            run.text += str(child)
                        else:
                            process_node(child, container) 
                else:
                    # Formatting tag outside a paragraph — create a new paragraph
                    para = container.add_paragraph()
                    process_node(node, para)
                processed_elements.add(node)

            elif node.name == 'div':
                block_children = node.find(['p', 'h1', 'h2', 'h3', 'h4', 'ul', 'ol', 'table', 'div'])
                if not block_children:
                    para = container.add_paragraph()
                    for child in node.contents:
                        process_node(child, para)
                    processed_elements.add(node)
                else:
                    for child in node.contents:
                        process_node(child, container)
            else:
                # Unknown tag — just recurse
                for child in node.contents:
                    process_node(child, container)

        # ── Start recursion ──
        for child in soup.contents:
            process_node(child, doc)

        styled_doc = self.style_manager.apply_template_styles(doc)
        self._apply_final_features(styled_doc)
        return styled_doc

    def _apply_final_features(self, doc: Document):
        """Adds Cover Page and Table of Contents if requested."""
        if self.include_cover:
            self.cover_manager.create_cover_page(doc)

        if self.include_toc:
            self.toc_manager.insert_toc(doc)
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # DOCX PIPELINE
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    def _pipeline_docx(self, path: str, out: str) -> dict:
        source_doc = Document(path)
        new_doc    = Document()
        self._h4_counters = {}  

        # ── STEP 1: EXTRACT ──────────────────────────────────────────────
        
        raw_lines = []

        for p_el in source_doc.element.iter(f'{{{W}}}p'):

            # Walk ancestor chain looking for Fallback tag
            is_fallback = False
            parent = p_el.getparent()
            while parent is not None:
                tag = parent.tag.split('}')[-1] if '}' in parent.tag else parent.tag
                if tag == 'Fallback':
                    is_fallback = True
                    break
                parent = parent.getparent()
            if is_fallback:
                continue

            para       = Paragraph(p_el, source_doc)
            style_name = para.style.name if para.style else 'Normal'

            # Word list-numbering: numId > 0 means this paragraph is part of
            # an auto-numbered list. The visible number is generated by Word's
            # engine — it is NOT stored in the text content.
            pPr      = p_el.find(qn('w:pPr'))
            numPr    = pPr.find(qn('w:numPr'))   if pPr   is not None else None
            numId_el = numPr.find(qn('w:numId')) if numPr is not None else None
            ilvl_el  = numPr.find(qn('w:ilvl'))  if numPr is not None else None
            num_id   = int(numId_el.get(qn('w:val'), 0)) if numId_el is not None else 0
            num_lvl  = int(ilvl_el.get(qn('w:val'),  0)) if ilvl_el  is not None else 0

            # Run-level signals: bold and explicit font size
            run_sizes, is_bold = [], False
            for r in p_el.findall(qn('w:r')):
                rPr = r.find(qn('w:rPr'))
                if rPr is not None:
                    sz = rPr.find(qn('w:sz'))
                    if sz is not None:
                        try: run_sizes.append(int(sz.get(qn('w:val'))))
                        except: pass
                    if rPr.find(qn('w:b')) is not None:
                        is_bold = True
            run_size = max(run_sizes) if run_sizes else 0

            text_full = para.text.strip()
            if not text_full:
                continue

            # Soft line-breaks (\v / \n) inside one paragraph → split into separate lines
            lines = text_full.split('\n') if '\n' in text_full else [text_full]
            for line in lines:
                line = line.strip()
                if line:
                    raw_lines.append({
                        'text':    line,
                        'style':   style_name,
                        'is_bold': is_bold,
                        'run_size': run_size,
                        'num_id':  num_id,
                        'num_lvl': num_lvl,
                    })

        # ── STEP 2: JOIN ─────────────────────────────────────────────────
        # Word COM PDF→DOCX conversion fragments numbered items across paragraphs:

        joined = []
        i = 0
        while i < len(raw_lines):
            item = raw_lines[i]
            t    = item['text'].strip()

            # Lone digit — attempt forward stitch
            if re.match(r'^\d+$', t) and i + 1 < len(raw_lines):
                next_t = raw_lines[i + 1]['text'].strip()

                if re.match(r'^[\.\)]\s+\S', next_t):
                    # "1" + ". Customer opens..." → "1. Customer opens..."
                    # Normalize both "." and ")" separators to ". "
                    normalized = re.sub(r'^[\.\)]\s*', '. ', next_t)
                    item = dict(item)
                    item['text'] = t + normalized
                    joined.append(item)
                    i += 2
                    continue

                elif next_t in ('.', ')') and i + 2 < len(raw_lines):
                    # "3" + ")" + ". System validates..." → "3. System validates..."
                    rest       = raw_lines[i + 2]['text'].strip()
                    rest_clean = re.sub(r'^[\.\)]\s*', '', rest)
                    item = dict(item)
                    item['text'] = t + '. ' + rest_clean
                    joined.append(item)
                    i += 3
                    continue

                elif re.match(r'^\d+$', next_t):
                    # Two consecutive lone digits — skip this one
                    i += 1
                    continue

                else:
                    # Lone digit with no joinable continuation — keep as placeholder
                    item = dict(item)
                    item['text'] = t + '.'
                    joined.append(item)
                    i += 1
                    continue

            # Lone punctuation — append to previous only if it isn't already a complete numbered item
            elif joined and len(t) == 1 and t in '.)':
                if not re.match(r'^\d+[\.\)]', joined[-1]['text']):
                    joined[-1]['text'] += t
                i += 1
                continue

            # Fragment like ". Customer scans QR on table"
            elif joined and re.match(r'^[\.\)]\s+\S', t):
                prev = joined[-1]['text']
                if not re.match(r'^\d+[\.\)]\s+\S', prev):
                    # Previous is not yet a complete numbered item — stitch
                    joined[-1]['text'] += t
                    i += 1
                    continue
                else:
                    # Previous is already a complete numbered item — this is a
                    # continuation bullet from the same step; strip leading dot
                    item = dict(item)
                    item['text'] = re.sub(r'^[\.\)]\s*', '', t).strip()
                    if item['text']:
                        joined.append(item)
                    i += 1
                    continue

            # Lone "-" paragraph — prefix the NEXT line as a dash bullet
            elif t == '-':
                if i + 1 < len(raw_lines):
                    raw_lines[i + 1]['text'] = '- ' + raw_lines[i + 1]['text'].lstrip('- ')
                i += 1
                continue

            if t:
                joined.append(item)
            i += 1

        raw_lines = joined

        # ── STEP 3: CLASSIFY + BUILD ──────────────────────────────────────
        
        expecting_list = False

        for item in raw_lines:
            text     = item['text'].replace('Ł', '').strip()
            style    = item['style']
            is_bold  = item['is_bold']
            run_size = item['run_size']
            num_id   = item['num_id']
            num_lvl  = item['num_lvl']
            is_numbered_list_item = num_id > 0

            if not text:
                continue

            # Structural pattern signals — zero keywords, work for any language/domain
            is_section_label    = text.endswith(':') and len(text.split()) <= 6
            is_numbered_heading = bool(re.match(r'^\d+[\.\)]\s+[A-Z]', text))
            is_explicit_bullet  = text.startswith(('☐', '☑', '•', '▪', '►', '■', '□'))
            is_dash_bullet      = bool(re.match(r'^[-–]\s+\S', text))

            # Update list context state
            if text.endswith('.') and len(text) > 80:
                expecting_list = False
            if is_section_label:
                expecting_list = True

            # ── H1 ───────────────────────────────────────────────────────
            # Word COM reliably preserves Heading 1 style name only
            if 'Heading 1' in style:
                new_doc.add_heading(text, level=1)
                expecting_list = False

            # ── H2 ───────────────────────────────────────────────────────
            
            elif (
                'Heading 2' in style
                or run_size >= 26
                or (is_bold and run_size >= 22 and len(text.split()) <= 10)
                or (is_bold and not is_section_label and len(text.split()) <= 8
                    and not text.endswith('.'))
                or re.match(r'^[A-Z][a-z]+\s+[\d][\d\-\+\.]*\s*:', text)
            ):
                new_doc.add_heading(text, level=2)
                expecting_list = False

            # ── H3 ───────────────────────────────────────────────────────
            # Section labels ("Tasks:", "Deliverables:") and short bold lines
            elif (
                'Heading 3' in style
                or is_section_label
                or (is_bold and len(text.split()) <= 6)
            ):
                new_doc.add_heading(text, level=3)
                expecting_list = True if is_section_label else False

            # ── H4 ───────────────────────────────────────────────────────
            # numId signal = Word auto-numbered list item ("1. Kickoff Meeting")
            # Short numbered heading = text already contains number ("2. UI/UX")
            # Guard: long text (>6 words) stays as a bullet — it's a step, not a title
            elif (
                (is_numbered_list_item and num_lvl == 0)
                or (is_numbered_heading and len(text.split()) <= 6)
            ):
                if is_numbered_list_item:
                    # Number is NOT in the text — Word generates it via numId.
                    # We reconstruct it manually so it appears in the output.
                    self._h4_counters[num_id] = self._h4_counters.get(num_id, 0) + 1
                    display_text = f"{self._h4_counters[num_id]}. {text}"
                else:
                    # Number already present in text (e.g. "2. UI/UX Requirements")
                    display_text = text
                new_doc.add_heading(display_text, level=4)
                expecting_list = True

            # ── BULLETS ──────────────────────────────────────────────────
            elif is_explicit_bullet or is_dash_bullet or expecting_list:
                clean = re.sub(r'^[-–☐☑•▪►■□]\s*', '', text).strip().rstrip('-').strip()
                if not clean:
                    continue

                # Code-style label heading: "FR-001: title" or "UC-01: description"
                # Detect by digit/hyphen in label — avoids promoting generic fields
                # like "Description:" or "User Role:" which have no digit/hyphen
                colon_pos = clean.find(':')
                label     = clean[:colon_pos] if colon_pos > 0 else ''
                is_label_heading = (
                    colon_pos > 0
                    and len(label.split()) <= 2
                    and not clean.endswith(':')
                    and len(clean) > colon_pos + 2
                    and not expecting_list
                    and re.search(r'[\d\-]', label)
                )

                if is_label_heading:
                    new_doc.add_heading(clean, level=3)
                else:
                    # Split crammed items — only when clearly 3+ distinct items (2+ words each)
                    # e.g. "Stakeholder map Interview schedule Project doc" → 3 bullets
                    items = re.split(r'(?<=[a-z\)\d])\s+(?=[A-Z][a-z])', clean)
                    items = [it.strip() for it in items if it.strip()]
                    if len(items) >= 3 and all(len(it.split()) >= 2 for it in items):
                        for it in items:
                            it_clean = re.sub(r'^[-☐•▪]\s*', '', it).strip()
                            if it_clean:
                                try:    p = new_doc.add_paragraph(style='List Bullet')
                                except: p = new_doc.add_paragraph()
                                p.add_run(it_clean)
                    else:
                        try:    p = new_doc.add_paragraph(style='List Bullet')
                        except: p = new_doc.add_paragraph()
                        p.add_run(clean)

            # ── BODY ─────────────────────────────────────────────────────
            else:
                new_doc.add_paragraph(text)

        # ── STEP 4: BRAND ────────────────────────────────────────────────
        branded_doc = self.style_manager.apply_template_styles(new_doc)
        
        # Apply optional features
        self._apply_final_features(branded_doc)
        
        branded_doc.save(out)
        import time
        time.sleep(1)
        return {'success': True, 'paragraphs': len(new_doc.paragraphs)}

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # PDF PIPELINE
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    def _pipeline_pdf(self, path: str, out: str) -> dict:
        """Adobe Extract API → semantic JSON elements → branded DOCX."""
        try:
            elements    = adobe_pdf_extract(path, ADOBE_CLIENT_ID, ADOBE_CLIENT_SECRET)
            doc         = self._build_from_adobe_json(elements)
            branded_doc = self.style_manager.apply_template_styles(doc)
            
            # Apply optional features
            self._apply_final_features(branded_doc)

            branded_doc.save(out)
            return {'success': True}
        except Exception as e:
            print(f'[DocumentProcessor] Adobe API error: {e}')
            return {'success': False, 'error': str(e)}

    def _build_from_adobe_json(self, elements: list) -> Document:
        """Convert Adobe Extract API semantic elements into a structured DOCX."""
        doc            = Document(self.template_path)
        expecting_list = False

        for el in elements:
            path = el.get('Path', '')
            text = el.get('Text', '').strip()

            if not text and 'Table' not in path:
                continue

            # Adobe separates bullet labels ("/Lbl") from content ("/LBody").
            # Skip labels — LBody carries the full text we need.
            if '/Lbl' in path:
                continue

            text = text.replace('Ł', '').strip()

            is_section_label = text.endswith(':') and len(text.split()) <= 6
            is_explicit_bullet = text.startswith(('-', '☐', '•', '▪', '➤', '■'))

            # Update list context
            if is_section_label:
                expecting_list = True
            elif text.endswith('.') or len(text) > 120 or re.search(r'/H\d', path):
                expecting_list = False

            if '/Title' in path:
                doc.add_heading(text, level=1)
            elif re.search(r'/H1', path):
                doc.add_heading(text, level=1)
            elif re.search(r'/H2', path):
                doc.add_heading(text, level=2)
            elif re.search(r'/H3', path) or is_section_label:
                doc.add_heading(text, level=3)
            elif '/LI' in path or '/LBody' in path or is_explicit_bullet or expecting_list:
                clean = re.sub(r'^[-\[\]☐•▪➤■]\s*', '', text)
                doc.add_paragraph(clean, style='List Bullet')
            elif 'Table' in path:
                pass  # table reconstruction not yet implemented
            else:
                doc.add_paragraph(text)

        return doc


    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # HTML TABLE HELPER
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    def _add_html_table(self, doc: Document, element):
        """Convert a BeautifulSoup <table> element into a Word table."""
        rows = element.find_all('tr')
        if not rows:
            return
        cols = max(len(r.find_all(['td', 'th'])) for r in rows)
        if not cols:
            return
        table = doc.add_table(rows=len(rows), cols=cols)
        try:    table.style = 'Table Grid'
        except Exception: pass
        for i, row in enumerate(rows):
            for j, cell in enumerate(row.find_all(['td', 'th'])):
                if j < cols:
                    table.rows[i].cells[j].text = cell.get_text(separator=' ').strip()
