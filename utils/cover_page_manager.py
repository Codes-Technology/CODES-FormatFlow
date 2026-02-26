"""
Cover page
"""

import ollama
import re
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from datetime import datetime


class CoverPageManager:

    def __init__(self, model_name: str = "phi3"):
        self.model_name = model_name
        self.bullet_prefixes = ('•', '-', '*', '▪', '►', '➤', '→', '○', '●', '‣', '–', '—')

    def create_cover_page(self, doc: Document) -> bool:
        title, subtitle = self._extract_title_subtitle(doc)
        if not title:
            print("[Cover Page] No title found, skipping")
            return False

        print(f"[Cover Page] Title: {title[:50]}")
        doc_text = self._extract_document_text(doc)
        summary = self._generate_summary_with_ollama(doc_text, title)

        original_content = self._store_document_content(doc)
        self._clear_document(doc)

        # Build cover content
        cover_paragraphs = self._build_cover_page(doc, title, subtitle, summary)

        # Add section break WITH vAlign=center (defines cover page section)
        self._add_cover_section_break(doc)

        # Restore all original content after the break
        skip_count = self._calculate_skip_count(original_content, title, subtitle)
        self._restore_content(doc, original_content, skip_count)

        # This prevents StyleManager from overriding the alignment
        self._force_center_on_cover(doc, len(cover_paragraphs))

        print("[Cover Page] ✅ Done!")
        return True

    # ================================================================
    # COVER CONTENT
    # ================================================================
    def _build_cover_page(self, doc: Document, title: str, subtitle: str, summary: str) -> list:
        """Build cover content. Returns list of added paragraphs."""
        added = []

        def add_centered(text, size, bold=False, italic=False):
            p = doc.add_paragraph()
            # Set alignment BOTH ways to be safe
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            self._force_para_center_xml(p)  # Also via XML
            p.paragraph_format.space_after = Pt(12)
            r = p.add_run(text)
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.italic = italic
            r.font.name = "Calibri"
            r.font.color.rgb = RGBColor(0, 0, 0)
            added.append(p)
            return p

        add_centered(title.upper(), 22, bold=True)

        if subtitle:
            add_centered(subtitle, 14)

        if summary:
            add_centered("Executive Summary", 13, bold=True)
            add_centered(summary, 11, italic=True)

        add_centered(datetime.now().strftime("%B %d, %Y"), 11)

        return added

    def _force_para_center_xml(self, para):
        """Force CENTER alignment directly in XML — cannot be overridden."""
        pPr = para._element.get_or_add_pPr()

        # Remove existing jc (justification) element
        for existing in pPr.findall(qn('w:jc')):
            pPr.remove(existing)

        # Add center justification
        jc = OxmlElement('w:jc')
        jc.set(qn('w:val'), 'center')
        pPr.append(jc)

    def _force_center_on_cover(self, doc: Document, cover_para_count: int):
        """
        Re-force CENTER alignment on cover paragraphs.
        Called LAST to prevent StyleManager from overriding.
        Cover paragraphs are the first N paragraphs in the document.
        """
        count = 0
        for para in doc.paragraphs:
            if count >= cover_para_count:
                break
            # Check if paragraph is part of cover (before section break)
            pPr = para._element.find(qn('w:pPr'))
            if pPr is not None:
                sectPr = pPr.find(qn('w:sectPr'))
                if sectPr is not None:
                    break  # This is the section break paragraph, stop

            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            self._force_para_center_xml(para)
            count += 1

    # ================================================================
    # SECTION BREAK WITH vAlign=center
    # ================================================================
    def _add_cover_section_break(self, doc: Document):
        """
        Inline sectPr defines the cover page section.
        vAlign=center here centers all content above this paragraph.
        Copies header/footer references from the LAST section (which has them).
        """
        from copy import deepcopy

        p = doc.add_paragraph()
        pPr = p._element.get_or_add_pPr()

        sectPr = OxmlElement('w:sectPr')

        # ── Copy header/footer references from doc's final sectPr ──
        # doc.sections[0]._sectPr is the LAST (final) section which has
        # the header/footer already set by _copy_header_footer()
        final_sectPr = doc.sections[0]._sectPr

        tags_to_copy = [
            'w:headerReference',
            'w:footerReference',
            'w:pgSz',
            'w:pgMar',
        ]

        for tag in tags_to_copy:
            for elem in final_sectPr.findall(qn(tag)):
                sectPr.append(deepcopy(elem))

        # Vertical center for cover page
        vAlign = OxmlElement('w:vAlign')
        vAlign.set(qn('w:val'), 'center')
        sectPr.append(vAlign)

        # Next page break
        sectType = OxmlElement('w:type')
        sectType.set(qn('w:val'), 'nextPage')
        sectPr.append(sectType)

        pPr.append(sectPr)
        print("[Cover Page] ✓ Header/footer references copied to cover section")

    # ================================================================
    # AI SUMMARY
    # ================================================================
    def _generate_summary_with_ollama(self, doc_text: str, title: str) -> str:
        if len(doc_text) > 6000:
            doc_text = doc_text[:6000]
        try:
            response = ollama.chat(
                model=self.model_name,
                messages=[{'role': 'user', 'content':
                    f"Summarize in 1-2 sentences:\nTitle: {title}\n\n{doc_text}"}],
                options={'temperature': 0.3, 'num_predict': 150}
            )
            summary = response['message']['content'].strip()
            for prefix in ["Executive Summary:", "Summary:", "Here is", "Here's"]:
                if summary.lower().startswith(prefix.lower()):
                    summary = summary[len(prefix):].strip()
            if summary.startswith('"'): summary = summary[1:]
            if summary.endswith('"'): summary = summary[:-1]
            if not summary.endswith('.'): summary += '.'
            sentences = re.split(r'(?<=[.!?])\s+', summary)
            return ' '.join(sentences[:2]) if len(sentences) > 2 else summary
        except Exception as e:
            print(f"[Cover Page] AI error: {e}")
            return "Comprehensive technical documentation and strategic recommendations."

    # ================================================================
    # TITLE EXTRACTION
    # ================================================================
    def _extract_title_subtitle(self, doc: Document) -> tuple:
        candidates = []
        for para in doc.paragraphs[:15]:
            text = para.text.strip()
            if not text or any(text.startswith(p) for p in self.bullet_prefixes):
                continue
            if re.match(r'^\d+\.', text):
                break
            if 2 <= len(text.split()) <= 25:
                candidates.append(text)
            if len(candidates) >= 2:
                break
        return (candidates[0] if candidates else None,
                candidates[1] if len(candidates) > 1 else None)

    def _extract_document_text(self, doc: Document) -> str:
        return '\n'.join([p.text.strip() for p in doc.paragraphs if p.text.strip()])

    # ================================================================
    # CONTENT STORAGE & RESTORATION
    # ================================================================
    def _store_document_content(self, doc: Document) -> list:
        content = []
        for para in doc.paragraphs:
            content.append({
                'text': para.text,
                'style_name': para.style.name if para.style else 'Normal',
                'alignment': para.alignment,
                'left_indent': para.paragraph_format.left_indent,
                'right_indent': para.paragraph_format.right_indent,
                'first_line_indent': para.paragraph_format.first_line_indent,
                'space_before': para.paragraph_format.space_before,
                'space_after': para.paragraph_format.space_after,
                'runs': [{
                    'text': r.text, 'bold': r.bold, 'italic': r.italic,
                    'underline': r.underline, 'font_size': r.font.size,
                    'font_name': r.font.name,
                    'font_color': r.font.color.rgb if r.font.color and r.font.color.rgb else None
                } for r in para.runs]
            })
        return content

    def _clear_document(self, doc: Document):
        for para in doc.paragraphs[:]:
            para._element.getparent().remove(para._element)

    def _calculate_skip_count(self, content: list, title: str, subtitle: str) -> int:
        skip = 0
        for item in content:
            text = item['text'].strip()
            if not text:
                skip += 1
            elif title and text == title:
                skip += 1
            elif subtitle and text == subtitle:
                skip += 1
            else:
                break
        return skip

    def _restore_content(self, doc: Document, content: list, skip_count: int = 0):
        for item in content[skip_count:]:
            p = doc.add_paragraph()
            try:
                p.style = doc.styles[item['style_name']]
            except:
                pass
            if item['alignment'] is not None:
                p.alignment = item['alignment']
            if item['left_indent']:
                p.paragraph_format.left_indent = item['left_indent']
            if item['right_indent']:
                p.paragraph_format.right_indent = item['right_indent']
            if item['first_line_indent']:
                p.paragraph_format.first_line_indent = item['first_line_indent']
            if item['space_before']:
                p.paragraph_format.space_before = item['space_before']
            if item['space_after']:
                p.paragraph_format.space_after = item['space_after']
            if item['runs']:
                for rd in item['runs']:
                    if rd['text']:
                        r = p.add_run(rd['text'])
                        if rd['bold'] is not None: r.bold = rd['bold']
                        if rd['italic'] is not None: r.italic = rd['italic']
                        if rd['underline'] is not None: r.underline = rd['underline']
                        if rd['font_size']: r.font.size = rd['font_size']
                        if rd['font_name']: r.font.name = rd['font_name']
                        if rd['font_color']: r.font.color.rgb = rd['font_color']
            elif item['text']:
                p.add_run(item['text'])
