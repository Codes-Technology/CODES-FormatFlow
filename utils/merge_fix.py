import copy

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement



def inject_numbering_part(doc: Document):
    try:
        _ = doc.part.numbering_part
        return
    except NotImplementedError:
        pass
    except Exception:
        pass

    from docx.opc.part import Part
    from docx.opc.packuri import PackURI

    numbering_xml = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<w:numbering xmlns:w="http://schemas.openxmlformats.org/'
        b'wordprocessingml/2006/main"></w:numbering>'
    )

    part = Part(
        PackURI('/word/numbering.xml'),
        'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
        numbering_xml,
        doc.part.package
    )
    doc.part.relate_to(
        part,
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering'
    )
    print("[Merge] Injected numbering part")


def safe_merge(master_doc: Document, sub_docs: list) -> Document:
    from docxcompose.composer import Composer
    inject_numbering_part(master_doc)
    composer = Composer(master_doc)

    for i, sub_doc in enumerate(sub_docs):
        try:
            inject_numbering_part(sub_doc)
            composer.append(sub_doc)
            print(f"[Merge] Doc {i+1} merged via docxcompose")
        except Exception as e:
            print(f"[Merge] docxcompose failed ({e}), using fallback...")
            _manual_append(master_doc, sub_doc)
    return master_doc


def _manual_append(master_doc: Document, sub_doc: Document):
    master_body = master_doc.element.body
    sub_body    = sub_doc.element.body
    pb = OxmlElement('w:p')
    r  = OxmlElement('w:r')
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    r.append(br); pb.append(r)
    master_body.append(pb)
    for element in sub_body:
        tag = element.tag.split('}')[-1] if '}' in element.tag else element.tag
        if tag == 'sectPr':
            continue
        master_body.append(copy.deepcopy(element))
    print("[Merge] Doc merged via manual copy")






