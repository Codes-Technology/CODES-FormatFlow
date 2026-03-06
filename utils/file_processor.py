"""FileProcessor — Merges already-branded DOCX files into one final document.

"""

from typing import List
from docx import Document
from docx.opc.part import Part
from docx.opc.packuri import PackURI


class FileProcessor:
    def __init__(self, output_folder: str):
        self.output_folder = output_folder

    def process_file(self, file_paths: List[str], output_path: str, **kwargs) -> bool:
        include_cover = kwargs.get('include_cover', False)
        include_toc   = kwargs.get('include_toc',   False)

        if not file_paths:
            print('FileProcessor] Error: no files provided.')
            return False

        try:
            # Ensure every file has a numbering part (prevents crash on List styles)
            for fp in file_paths:
                try:
                    d = Document(fp)
                    _inject_numbering_if_missing(d)
                    d.save(fp)
                except Exception as e:
                    print(f'FileProcessor] Numbering inject warning ({fp}): {e}')

            # process file
            processed = Document(file_paths[0])

            # Cover FIRST (inserts at 0 → becomes page 1)
            if include_cover:
                try:
                    from utils.cover_page_manager import CoverPageManager
                    CoverPageManager().create_cover_page(processed)
                    print('FileProcessor] Cover page added.')
                except ImportError:
                    print('FileProcessor] CoverPageManager not found.')
                except Exception as e:
                    print(f'FileProcessor] Cover page failed: {e}')

            # TOC SECOND (inserts at 0 → appears after cover, before content)
            if include_toc:
                try:
                    from utils.toc_manager import TocManager
                    TocManager().insert_toc(processed)
                    print('FileProcessor] TOC added.')
                except ImportError:
                    print('FileProcessor] TocManager not found.')
                except Exception as e:
                    print(f'[FileProcessor] TOC failed: {e}')

            processed.save(output_path)
            print(f'[FileProcessor] ✓ Saved: {output_path}')
            return True

        except Exception as e:
            import traceback
            print(f'[FileProcessor] ✗ Critical: {e}')
            traceback.print_exc()
            return False


def _inject_numbering_if_missing(doc: Document):
    """Injects an empty numbering.xml part if the document has none."""
    try:
        _ = doc.part.numbering_part
        return
    except (NotImplementedError, AttributeError, Exception):
        pass

    try:
        xml = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               b'<w:numbering xmlns:w="http://schemas.openxmlformats.org/'
               b'wordprocessingml/2006/main"></w:numbering>')
        part = Part(
            PackURI('/word/numbering.xml'),
            'application/vnd.openxmlformats-officedocument'
            '.wordprocessingml.numbering+xml',
            xml, doc.part.package
        )
        doc.part.relate_to(
            part,
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering'
        )
    except Exception as e:
        print(f'[FileProcessor] Could not inject numbering: {e}')