import xml.etree.ElementTree as ET
import re
import sys
sys.stdout.reconfigure(encoding='utf-8')

# Namespaces
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
VML = 'urn:schemas-microsoft-com:vml'

ns = {'w': W, 'r': R, 'a': A, 'v': VML}

BASE = r'C:\Users\Daddy\Apps\TheRoaring20s\docx_extract\unpacked\word'

# Parse relationships: rId -> media filename
rels = {}
rels_tree = ET.parse(BASE + r'\_rels\document.xml.rels')
for rel in rels_tree.getroot():
    rid = rel.get('Id')
    target = rel.get('Target', '')
    rel_type = rel.get('Type', '')
    if 'image' in rel_type:
        rels[rid] = target  # e.g. "media/image5.jpeg"

# Parse document body
doc_tree = ET.parse(BASE + r'\document.xml')
body = doc_tree.getroot().find(f'.//{{{W}}}body')

def get_para_text(para):
    parts = []
    for t in para.iter(f'{{{W}}}t'):
        parts.append(t.text or '')
    return ''.join(parts).strip()

def get_images_in_para(para):
    images = []
    for blip in para.iter(f'{{{A}}}blip'):
        rid = blip.get(f'{{{R}}}embed')
        if rid and rid in rels:
            images.append(rels[rid])
    for imagedata in para.iter(f'{{{VML}}}imagedata'):
        rid = imagedata.get(f'{{{R}}}id')
        if rid and rid in rels:
            images.append(rels[rid])
    return images

# Walk all top-level elements, collecting vignettes
vignettes = []
current = None

for elem in body:
    tag = elem.tag

    # Handle paragraphs
    if tag == f'{{{W}}}p':
        text = get_para_text(elem)
        images = get_images_in_para(elem)

        # Check for VIGNETTE_START
        m = re.match(r'VIGNETTE_START\s*[:\-]?\s*(.+)', text, re.IGNORECASE)
        if m:
            label = m.group(1).strip()
            current = {'label': label, 'paragraphs': [], 'images': []}
            vignettes.append(current)
            continue

        # Check for VIGNETTE_END
        if re.match(r'VIGNETTE_END', text, re.IGNORECASE):
            current = None
            continue

        if current is not None:
            if text:
                current['paragraphs'].append(text)
            current['images'].extend(images)

    # Handle tables (tbl) - may contain text/images too
    elif tag == f'{{{W}}}tbl':
        if current is not None:
            for cell_para in elem.iter(f'{{{W}}}p'):
                text = get_para_text(cell_para)
                images = get_images_in_para(cell_para)
                if text:
                    current['paragraphs'].append(text)
                current['images'].extend(images)

# Print manifest
print(f"Total vignettes found: {len(vignettes)}\n")
print("=" * 60)
for i, v in enumerate(vignettes):
    print(f"\nVIGNETTE {i+1}: {v['label']}")
    print(f"  Images ({len(v['images'])}):")
    for img in v['images']:
        print(f"    {img}")
    # Show first 80 chars of first paragraph as preview
    if v['paragraphs']:
        preview = v['paragraphs'][0][:80]
        print(f"  Text preview: {preview}...")
    else:
        print(f"  Text preview: (none)")
print("\n" + "=" * 60)

# Also list any images in doc NOT captured by any vignette
all_vignette_images = set()
for v in vignettes:
    for img in v['images']:
        all_vignette_images.add(img)

all_doc_images = set(rels.values())
uncaptured = all_doc_images - all_vignette_images
print(f"\nImages in document NOT inside any vignette ({len(uncaptured)}):")
for img in sorted(uncaptured):
    print(f"  {img}")
