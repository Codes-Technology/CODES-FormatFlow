"""
StyleManager 
"""

from docx import Document
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from copy import deepcopy

# Sizes in half-points
H1_SIZE   = 32   # 16pt
H2_SIZE   = 28   # 14pt
H3_SIZE   = 26   # 13pt
H4_SIZE   = 24   # 12pt
BODY_SIZE = 22   # 11pt
MIN_SIZE  = 22   # 11pt minimum

BODY_FONT  = 'Calibri'
BODY_COLOR = '000000'

# Relationship namespace
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


class StyleManager:
    def __init__(self, template_path: str):
        self.template_path = template_path
        self.template = Document(template_path)

    def apply_template_styles(self, doc: Document) -> Document:
        """Single entry point — full branding pipeline."""
        print('[StyleManager] Applying branding...')
        self._remove_leading_empty_paragraphs(doc)
        self._override_style_sizes(doc)
        self._fix_list_indentation(doc)
        self._brand_body_paragraphs(doc)
        self._apply_page_layout_rules(doc)
        self._apply_heading_indents(doc)
        self._copy_header_footer(doc)
        self._add_page_number(doc)
       

        print('[StyleManager] ✓ Done')
        return doc

    
    # STEP 1 — Style-level sizes (styles.xml, not run-by-run)
    

    def _override_style_sizes(self, doc: Document):
        config = {
            'Normal':       {'sz': BODY_SIZE, 'bold': False},
            # Paragraph heading styles
            'Heading1':     {'sz': H1_SIZE,   'bold': True},
            'Heading 1':    {'sz': H1_SIZE,   'bold': True},
            'Heading2':     {'sz': H2_SIZE,   'bold': True},
            'Heading 2':    {'sz': H2_SIZE,   'bold': True},
            'Heading3':     {'sz': H3_SIZE,   'bold': True},
            'Heading 3':    {'sz': H3_SIZE,   'bold': True},
            # Linked CHARACTER styles — Word renders heading text through these.
            # They override paragraph style size/color, so must match targets.
            'Heading1Char': {'sz': H1_SIZE,   'bold': True},
            'Heading2Char': {'sz': H2_SIZE,   'bold': True},
            'Heading3Char': {'sz': H3_SIZE,   'bold': True},
            'Heading4':     {'sz': 24,         'bold': True},
            'Heading 4':    {'sz': 24,         'bold': True},
            'Heading4Char': {'sz': 24,         'bold': True},
        }

        styles_elem = doc.part.styles._element
        for style_el in styles_elem.findall(qn('w:style')):
            style_id   = style_el.get(qn('w:styleId'), '')
            name_el    = style_el.find(qn('w:name'))
            style_name = name_el.get(qn('w:val'), '') if name_el is not None else ''

            cfg = config.get(style_id) or config.get(style_name)
            if cfg is None:
                continue

            rPr = style_el.find(qn('w:rPr'))
            if rPr is None:
                rPr = OxmlElement('w:rPr')
                style_el.append(rPr)

            for t in ('w:sz', 'w:szCs'):
                for e in rPr.findall(qn(t)):
                    rPr.remove(e)
            sz = OxmlElement('w:sz');   sz.set(qn('w:val'), str(cfg['sz']));   rPr.append(sz)
            szCs = OxmlElement('w:szCs'); szCs.set(qn('w:val'), str(cfg['sz'])); rPr.append(szCs)

            for e in rPr.findall(qn('w:color')): rPr.remove(e)
            c = OxmlElement('w:color'); c.set(qn('w:val'), BODY_COLOR); rPr.append(c)

            for e in rPr.findall(qn('w:rFonts')): rPr.remove(e)
            f = OxmlElement('w:rFonts')
            f.set(qn('w:ascii'), BODY_FONT); f.set(qn('w:hAnsi'), BODY_FONT)
            f.set(qn('w:cs'), BODY_FONT);   rPr.append(f)

            for e in rPr.findall(qn('w:b')): rPr.remove(e)
            if cfg['bold']:
                rPr.append(OxmlElement('w:b'))

            print(f'  [StyleManager] Style "{style_id or style_name}" → '
                  f'{cfg["sz"] / 2}pt bold={cfg["bold"]}')
            
            pPr = style_el.find(qn('w:pPr'))
            if pPr is None:
                pPr = OxmlElement('w:pPr')
                style_el.append(pPr)

            for e in pPr.findall(qn('w:spacing')):
                pPr.remove(e)

            spacing = OxmlElement('w:spacing')

            if cfg['bold']:
                spacing.set(qn('w:before'), '0')
                spacing.set(qn('w:after'), '200')
            else:
                spacing.set(qn('w:before'), '0')
                spacing.set(qn('w:after'), '120')

            pPr.append(spacing)

            
            if cfg['bold']:   
                for e in pPr.findall(qn('w:keepNext')):
                    pPr.remove(e)
                pPr.append(OxmlElement('w:keepNext'))

    def _apply_heading_indents(self, doc: Document):
        """Indent headings progressively right by level."""
        INDENT = {
            'Heading 1': 0,   'Heading1': 0,
            'Heading 2': 240, 'Heading2': 240,   # ~0.17 inch
            'Heading 3': 480, 'Heading3': 480,   # ~0.33 inch
            'Heading 4': 720, 'Heading4': 720,   # ~0.50 inch
        }
        for para in doc.paragraphs:
            sname = para.style.name if para.style else ''
            twips = INDENT.get(sname)
            if twips is None:
                continue
            pPr = para._p.find(qn('w:pPr'))
            if pPr is None:
                pPr = OxmlElement('w:pPr')
                para._p.insert(0, pPr)
            for e in pPr.findall(qn('w:ind')):
                pPr.remove(e)
            if twips > 0:
                ind = OxmlElement('w:ind')
                ind.set(qn('w:left'), str(twips))
                pPr.append(ind)
    
    def _fix_list_indentation(self, doc: Document):
        """Bullets under H4 indent further than bullets under H3."""

        # Base indent for all bullets
        BASE_LEFT    = 720   # twips (~0.5 inch)
        BASE_HANGING = 240   # twips (~0.17 inch)
        
        # Extra indent when bullet follows H4
        H4_LEFT      = 960   # twips (~0.67 inch)

        last_heading_level = 0

        for para in doc.paragraphs:
            sname = para.style.name if para.style else ''

            # Track heading level
            if sname in ('Heading 1', 'Heading1'):
                last_heading_level = 1
            elif sname in ('Heading 2', 'Heading2'):
                last_heading_level = 2
            elif sname in ('Heading 3', 'Heading3'):
                last_heading_level = 3
            elif sname in ('Heading 4', 'Heading4'):
                last_heading_level = 4

            elif 'List' in sname:
                left = H4_LEFT if last_heading_level == 4 else BASE_LEFT

                pPr = para._p.find(qn('w:pPr'))
                if pPr is None:
                    pPr = OxmlElement('w:pPr')
                    para._p.insert(0, pPr)

                for e in pPr.findall(qn('w:ind')):
                    pPr.remove(e)

                ind = OxmlElement('w:ind')
                ind.set(qn('w:left'),    str(left))
                ind.set(qn('w:hanging'), str(BASE_HANGING))
                pPr.append(ind)

    def _copy_header_footer(self, doc: Document):
        template_section = self.template.sections[0]
        for section in doc.sections:
            section.header.is_linked_to_previous = False
            section.footer.is_linked_to_previous = False
            self._deep_copy_hdr_ftr(template_section.header, section.header)
            self._deep_copy_hdr_ftr(template_section.footer, section.footer)
    
    def _add_page_number(self, doc: Document):
        
        for section in doc.sections:
            footer = section.footer

            p = footer.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

            run = p.add_run()

            fldChar1 = OxmlElement('w:fldChar')
            fldChar1.set(qn('w:fldCharType'), 'begin')
            run._r.append(fldChar1)

            instrText = OxmlElement('w:instrText')
            instrText.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            instrText.text = " PAGE "
            run._r.append(instrText)

            fldChar2 = OxmlElement('w:fldChar')
            fldChar2.set(qn('w:fldCharType'), 'end')
            run._r.append(fldChar2)

    def _deep_copy_hdr_ftr(self, source, target):
        rId_map = {}
        for old_rId, rel in source.part.rels.items():
            try:
                if 'image' in rel.reltype:
                    new_rId = target.part.relate_to(rel.target_part, rel.reltype)
                    rId_map[old_rId] = new_rId
                elif 'hyperlink' in rel.reltype:
                    new_rId = target.part.relate_to(
                        rel.target_ref, rel.reltype, is_external=True)
                    rId_map[old_rId] = new_rId
            except Exception as e:
                print(f'  [StyleManager] Rel copy warning ({old_rId}): {e}')

        target_elem = target._element
        for p in list(target_elem.findall(qn('w:p'))):
            target_elem.remove(p)

        for src_p in source._element.findall(qn('w:p')):
            new_p = deepcopy(src_p)
            if rId_map:
                _remap_rids(new_p, rId_map)
            target_elem.append(new_p)
    
    def _iter_all_paragraphs(self, doc: Document):
        """
        Yield every paragraph in the document,
        including paragraphs inside tables.
        """
        # Normal paragraphs
        for para in doc.paragraphs:
            yield para

        # Table paragraphs
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        yield para

    def _brand_body_paragraphs(self, doc: Document):
        heading_names = {
            'Heading 1', 'Heading 2', 'Heading 3', 'Heading 4',
            'Heading1',  'Heading2',  'Heading3',  'Heading4',
        }

        for para in self._iter_all_paragraphs(doc):
            style_name  = para.style.name if para.style else 'Normal'
            is_heading  = style_name in heading_names or style_name.startswith('Heading')

            for r in para._p.findall(qn('w:r')):
                rPr = r.find(qn('w:rPr'))
                if rPr is None:
                    rPr = OxmlElement('w:rPr')
                    r.insert(0, rPr)

                for e in rPr.findall(qn('w:color')): rPr.remove(e)
                ce = OxmlElement('w:color'); ce.set(qn('w:val'), BODY_COLOR); rPr.append(ce)

                for e in rPr.findall(qn('w:rFonts')): rPr.remove(e)
                fe = OxmlElement('w:rFonts')
                fe.set(qn('w:ascii'), BODY_FONT); fe.set(qn('w:hAnsi'), BODY_FONT)
                rPr.append(fe)

                if is_heading:
                    for e in rPr.findall(qn('w:sz')):   rPr.remove(e)
                    for e in rPr.findall(qn('w:szCs')): rPr.remove(e)
                else:
                    for e in rPr.findall(qn('w:b')): rPr.remove(e)
                    for e in rPr.findall(qn('w:i')): rPr.remove(e)
                    sz_el   = rPr.find(qn('w:sz'))
                    current = int(sz_el.get(qn('w:val'), '0')) if sz_el is not None else 0
                    if current < MIN_SIZE or current == 0:
                        for e in rPr.findall(qn('w:sz')):   rPr.remove(e)
                        for e in rPr.findall(qn('w:szCs')): rPr.remove(e)
                        s = OxmlElement('w:sz');   s.set(qn('w:val'), str(BODY_SIZE));   rPr.append(s)
                        sc = OxmlElement('w:szCs'); sc.set(qn('w:val'), str(BODY_SIZE)); rPr.append(sc)

    def _apply_page_layout_rules(self, doc: Document):
        first_h1 = True
        prev_style = None

        for para in doc.paragraphs:
            if not para.style:
                continue

            style_name = para.style.name

            if style_name in ('Heading 1', 'Heading1'):
                if not first_h1:
                    
                    if prev_style and not prev_style.startswith('Heading'):
                        para.paragraph_format.page_break_before = True
                first_h1 = False

            if style_name.startswith('Heading'):
                para.paragraph_format.keep_with_next = True
                pPr = para._p.find(qn('w:pPr'))
                if pPr is None:
                    pPr = OxmlElement('w:pPr')
                    para._p.insert(0, pPr)
                if pPr.find(qn('w:keepNext')) is None:
                    pPr.append(OxmlElement('w:keepNext'))

            if para.text.strip():
                prev_style = style_name
    
    def _remove_leading_empty_paragraphs(self, doc):
        while doc.paragraphs and doc.paragraphs[0].text.strip() == "":
            p = doc.paragraphs[0]._element
            p.getparent().remove(p)
    

def _remap_rids(element, rId_map: dict):
    
    for attr_key in list(element.attrib.keys()):
        local = attr_key.split('}')[-1] if '}' in attr_key else attr_key
        ns    = attr_key.split('}')[0].lstrip('{') if '}' in attr_key else ''

        if ns == R_NS and local in ('embed', 'id', 'link', 'href', 'pict'):
            old_val = element.attrib[attr_key]
            if old_val in rId_map:
                element.attrib[attr_key] = rId_map[old_val]

    for child in element:
        _remap_rids(child, rId_map)