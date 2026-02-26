import re
from statistics import mode, StatisticsError
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
try:
    from utils.cover_page_manager import CoverPageManager
except ImportError:
    CoverPageManager = None

class StyleManager:
    """Manages document styling, including headings, tables, and cover pages."""
    
    def __init__(self, template_path):
        self.template_path = template_path
        self.cover_page_manager = CoverPageManager() if CoverPageManager else None
        self.bullet_chars = ('•', '-', '*', '▪', '►', '➤', '→', '○', '●', '‣', '–', '—')

    def apply_template_styles(self, doc: Document) -> Document:
        print("[Style Manager] Applying template styles...")
        self._apply_base_styles(doc)
        self.reconstruct_tables(doc)
        self._apply_heading_styles(doc)
        self._apply_page_break_control(doc)
        
        if self.cover_page_manager:
            self.cover_page_manager.create_cover_page(doc)
            
        return doc

    def add_smart_heading(self, doc: Document, text: str, context_idx: int):
        level = 1
        if re.match(r'^\d+\.\d+', text): level = 2
        elif re.match(r'^\d+\.', text): level = 1
        elif text.isupper() and len(text.split()) < 10: level = 1
        
        p = doc.add_paragraph(text)
        p.style = f'Heading {level}'
        self._format_heading_run(p, level)

    def _apply_base_styles(self, doc: Document):
        """Ensure all text uses Calibri and is black."""
        for para in doc.paragraphs:
            if not para.style.name.startswith('Heading'):
                for run in para.runs:
                    run.font.name = 'Calibri'
                    run.font.color.rgb = RGBColor(0, 0, 0)
                    if run.font.size is None:
                        run.font.size = Pt(11)

    def _apply_heading_styles(self, doc: Document):
        count = 0
        for para in doc.paragraphs:
            text = para.text.strip()
            if not text or any(text.startswith(b) for b in self.bullet_chars):
                continue
            
            level = self._detect_heading_level(para, text)
            if level:
                try: para.style = doc.styles[f'Heading {level}']
                except KeyError: pass 
                self._format_heading_run(para, level)
                count += 1
        print(f"[Style Manager] Styled {count} headings")

    def _detect_heading_level(self, para, text: str):
        if 'Heading 1' in para.style.name or para.style.name == 'Title': return 1
        if 'Heading 2' in para.style.name or 'Heading 3' in para.style.name: return 2
        
        if re.match(r'^\d+\.\s+[A-Z]', text): return 1
        if re.match(r'^\d+\.\d+[\.\s]+[A-Z]', text): return 2
        if text.isupper() and 3 <= len(text.split()) <= 8: return 1
        
        runs_bold = [r for r in para.runs if r.bold and r.text.strip()]
        if runs_bold and all(r.bold for r in para.runs if r.text.strip()) and len(text) < 80 and len(text.split()) <= 10:
            return 2

        if (len(text) < 60 and len(text.split()) >= 2 and text[0].isupper()
            and text[-1] not in ('.', ',', ';', ':', '?', '!')
            and len(text.split()) <= 8):
            return 2
        return None

    def _format_heading_run(self, para, level: int):
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        para.paragraph_format.space_before = Pt(12)
        para.paragraph_format.space_after = Pt(6)
        
        for run in para.runs:
            run.font.bold = True
            run.font.name = "Calibri"
            run.font.color.rgb = RGBColor(0, 0, 0)
            run.font.size = Pt(16) if level == 1 else Pt(14)

    def reconstruct_tables(self, doc: Document):
        candidates = []
        for i, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            if not text or any(text.startswith(b) for b in self.bullet_chars) or re.match(r'^\d+\.', text):
                continue
                
            cols = re.split(r'\t+|\s{2,}', text)
            if len(cols) >= 2:
                candidates.append((i, para, cols))
                
        if not candidates: return

        groups = []
        if candidates:
            current_group = [candidates[0]]
            for k in range(1, len(candidates)):
                curr_idx = candidates[k][0]
                prev_idx = candidates[k-1][0]
                
                if curr_idx == prev_idx + 1:
                    current_group.append(candidates[k])
                else:
                    groups.append(current_group)
                    current_group = [candidates[k]]
            groups.append(current_group)

        tables_created = 0
        for group in reversed(groups):
            if len(group) < 2: continue
            
            col_counts = [len(c[2]) for c in group]
            try: most_common_cols = mode(col_counts)
            except StatisticsError: most_common_cols = max(set(col_counts), key=col_counts.count)
                
            if most_common_cols < 2: continue
                
            consistency_count = sum(1 for c in col_counts if c == most_common_cols)
            if consistency_count / len(group) < 0.8: continue 
                
            table = doc.add_table(rows=len(group), cols=most_common_cols)
            table.style = 'Table Grid'
            
            for row_idx, (_, _, cols) in enumerate(group):
                row_cells = table.rows[row_idx].cells
                for c_idx in range(min(len(cols), most_common_cols)):
                    row_cells[c_idx].text = cols[c_idx].strip()
                    
            # In-place table replacement
            first_para = group[0][1] 
            tbl_element = table._element
            p_element = first_para._element
            
            if tbl_element.getparent() is not None:
                tbl_element.getparent().remove(tbl_element)
            
            p_element.addprevious(tbl_element)
            
            for _, para, _ in group:
                p = para._element
                if p.getparent() is not None:
                    p.getparent().remove(p)
            
            tables_created += 1
            
        if tables_created > 0:
            print(f"[Style Manager] Reconstructed {tables_created} tables perfectly in-place")

    def _apply_page_break_control(self, doc: Document):
        for para in doc.paragraphs:
            if 'Heading' in para.style.name:
                para.paragraph_format.keep_with_next = True
                para.paragraph_format.keep_together = True
            else:
                para.paragraph_format.widow_control = True