"""
TocManager — Word-native Table of Contents

"""

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import RGBColor


class TocManager:

    def insert_toc(self, doc: Document):
        """Insert a Word-native TOC at the top of the document."""
        self._make_toc_black(doc)

        title_para = self._make_title_para('Table of Contents')

        # ── Build the TOC field paragraph ──
        # Each fldChar must be in its OWN <w:r> — Word spec requires this.
        toc_para = OxmlElement('w:p')

        def _rpr():
            rPr = OxmlElement('w:rPr')
            rPr.append(OxmlElement('w:noProof'))
            return rPr

        # Run 1: BEGIN — dirty=true forces recalculation every time doc opens
        r1 = OxmlElement('w:r')
        r1.append(_rpr())
        fc_begin = OxmlElement('w:fldChar')
        fc_begin.set(qn('w:fldCharType'), 'begin')
        fc_begin.set(qn('w:dirty'), 'true')
        r1.append(fc_begin)
        toc_para.append(r1)

        # Run 2: INSTRUCTION
        r2 = OxmlElement('w:r')
        r2.append(_rpr())
        instrText = OxmlElement('w:instrText')
        instrText.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        instrText.text = ' TOC \\o "1-3" \\h \\z '
        r2.append(instrText)
        toc_para.append(r2)

        # Run 3: SEPARATE — required, divides instruction from placeholder
        r3 = OxmlElement('w:r')
        r3.append(_rpr())
        fc_sep = OxmlElement('w:fldChar')
        fc_sep.set(qn('w:fldCharType'), 'separate')
        r3.append(fc_sep)
        toc_para.append(r3)

        # Run 4: placeholder text (Word replaces this on open)
        r4 = OxmlElement('w:r')
        r4.append(_rpr())
        t_ph = OxmlElement('w:t')
        t_ph.text = 'Table of contents will appear here when opened in Word.'
        r4.append(t_ph)
        toc_para.append(r4)

        # Run 5: END
        r5 = OxmlElement('w:r')
        r5.append(_rpr())
        fc_end = OxmlElement('w:fldChar')
        fc_end.set(qn('w:fldCharType'), 'end')
        r5.append(fc_end)
        toc_para.append(r5)

        # ── Insert at top of document ──
        # Insert in reverse order: each insert(0) pushes the previous down.
        # Final order: title → toc_para → page_break → original content
        page_break = self._make_page_break()
        body = doc.element.body
        body.insert(0, page_break)   # inserted 3rd → sits between TOC and content
        body.insert(0, toc_para)     # inserted 2nd → index 1
        body.insert(0, title_para)   # inserted 1st → index 0

        self._enable_auto_update(doc)
        print('[TocManager] ✓ Native Word TOC inserted')

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # TITLE PARAGRAPH
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    def _make_title_para(self, title: str):
        """Build the 'Table of Contents' heading paragraph as raw XML."""
        p   = OxmlElement('w:p')
        pPr = OxmlElement('w:pPr')

        jc = OxmlElement('w:jc')
        jc.set(qn('w:val'), 'center')
        pPr.append(jc)

        sp = OxmlElement('w:spacing')
        sp.set(qn('w:after'), '240')
        pPr.append(sp)

        p.append(pPr)

        r   = OxmlElement('w:r')
        rPr = OxmlElement('w:rPr')
        rPr.append(OxmlElement('w:b'))

        color = OxmlElement('w:color')
        color.set(qn('w:val'), '000000')
        rPr.append(color)

        fonts = OxmlElement('w:rFonts')
        fonts.set(qn('w:ascii'), 'Calibri')
        fonts.set(qn('w:hAnsi'), 'Calibri')
        rPr.append(fonts)

        sz = OxmlElement('w:sz')
        sz.set(qn('w:val'), '32')   # 16pt
        rPr.append(sz)

        r.append(rPr)
        t = OxmlElement('w:t')
        t.text = title
        r.append(t)
        p.append(r)
        return p

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # TOC STYLES — black, no underline
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    def _make_toc_black(self, doc: Document):
        """
        Force TOC 1/2/3 styles to black with no underline.

        IMPORTANT: style.font.color.rgb = None does NOT set black.
        It CLEARS the color, so it inherits from the Hyperlink character
        style which is blue + underlined. Must use RGBColor(0, 0, 0).
        Also patch XML directly to defeat Hyperlink style inheritance.
        """
        for style_name in ('TOC 1', 'TOC 2', 'TOC 3'):
            try:
                style = doc.styles[style_name]
            except KeyError:
                style = doc.styles.add_style(style_name, 1)

            style.font.color.rgb = RGBColor(0, 0, 0)
            style.font.underline = False
            style.font.name = 'Calibri'

            # Also write directly to XML to defeat Hyperlink style inheritance
            style_elem = style.element
            rPr = style_elem.find(qn('w:rPr'))
            if rPr is None:
                rPr = OxmlElement('w:rPr')
                style_elem.append(rPr)

            for c in rPr.findall(qn('w:color')): rPr.remove(c)
            color_el = OxmlElement('w:color')
            color_el.set(qn('w:val'), '000000')
            rPr.append(color_el)

            for u in rPr.findall(qn('w:u')): rPr.remove(u)
            u_el = OxmlElement('w:u')
            u_el.set(qn('w:val'), 'none')
            rPr.append(u_el)

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # PAGE BREAK
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    def _make_page_break(self):
        """
        Simple paragraph containing a page-break run.
        Keeps the document in ONE section so TOC page numbers are accurate.
        See module docstring for why section breaks must not be used here.
        """
        p  = OxmlElement('w:p')
        r  = OxmlElement('w:r')
        br = OxmlElement('w:br')
        br.set(qn('w:type'), 'page')
        r.append(br)
        p.append(r)
        return p

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # AUTO UPDATE
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    def _enable_auto_update(self, doc: Document):
        """Tell Word to update all fields (including TOC) on open."""
        try:
            settings = doc.settings.element
            for existing in settings.findall(qn('w:updateFields')):
                settings.remove(existing)
            uf = OxmlElement('w:updateFields')
            uf.set(qn('w:val'), 'true')
            settings.append(uf)
        except Exception:
            pass