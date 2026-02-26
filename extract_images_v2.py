
import os
from docx import Document

def extract_images():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    docx_path = os.path.join(base_dir, "Letter_pad.docx")
    
    if not os.path.exists(docx_path):
        print(f"File not found: {docx_path}")
        return

    doc = Document(docx_path)
    
    def save_image_from_part(part, prefix):
        try:
            blob = part.blob
            # Simple hash to identify duplicates if needed, but for now just sequential
            # We need to know which is which.
            # Let's verify image size/dimensions if possible?
            # Instead, just save them all.
            pass
        except: pass

      
    section = doc.sections[0]
    header = section.header
    
    print("--- Header Images ---")
    header_rels = sorted(header.part.rels.items(), key=lambda x: x[0]) # Start with rId
    
    # We want to map rId to filename
    rId_map = {}
    
    # First pass: save all images referenced
    img_count = 0
    for rId, rel in header_rels:
        if "image" in rel.target_ref:
            img_count += 1
            fname = f"header_img_{img_count}.png"
            blob = rel.target_part.blob
            with open(os.path.join(base_dir, fname), 'wb') as f:
                f.write(blob)
            print(f"Saved {fname} (rId: {rId}, Size: {len(blob)})")
            rId_map[rId] = fname

    # Second pass: Inspect XML to see which rId is in which paragraph
    # This matches the images to their visual order
    for i, p in enumerate(header.paragraphs):
        for run in p.runs:
            # Find blip
            # qn('a:blip') is hard because findAll expects xpath or tag name
            # namespaces={'a': 'http://schemas.openxmlformats.org/drawingml/2006/main', 'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
            
            blips = run._element.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
            for blip in blips:
                embed = blip.get(f"{{http://schemas.openxmlformats.org/officeDocument/2006/relationships}}embed")
                if embed in rId_map:
                    print(f"Para {i} uses {rId_map[embed]}")

if __name__ == "__main__":
    extract_images()
