import xml.etree.ElementTree as ET
import re
import os
import shutil
import json
import sys
sys.stdout.reconfigure(encoding='utf-8')

# ── namespaces ────────────────────────────────────────────────────────────────
W   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
A   = 'http://schemas.openxmlformats.org/drawingml/2006/main'
VML = 'urn:schemas-microsoft-com:vml'

BASE   = r'C:\Users\Daddy\Apps\TheRoaring20s\docx_extract\unpacked\word'
ROOT   = r'C:\Users\Daddy\Apps\TheRoaring20s'
VIG_DIR = os.path.join(ROOT, 'vignettes')
IMG_DIR = os.path.join(ROOT, 'images')
MEDIA   = os.path.join(BASE, 'media')

os.makedirs(VIG_DIR, exist_ok=True)
os.makedirs(IMG_DIR, exist_ok=True)

# ── vignette → country mapping (in document order) ───────────────────────────
VIGNETTE_MAP = [
    ('FRANCE PARIS',                    'france',    1),
    ('CHINA SHANGHAI',                  'china',     1),
    ('GERMANY WOZZECK',                 'germany',   1),
    ('CHINA Shanghai_2',                'china',     2),
    ('Shanghai 3',                      'china',     3),
    ('GERMANY THREEPENNY KITKAT LATENIGHT', 'germany', 2),
    ('USA HARLEM RENT PARTY',           'us',        1),
    ('USA TIN PAN ALLEY',               'us',        2),
    ('USA HARLEM GERSHWIN AND THE DUKE','us',        3),
    ('USA HARLEM HI DI HO',             'us',        4),
    ('AGENTINA TANGO',                  'argentina', 1),
    ('EGYPT Cairo',                     'egypt',     1),
    ('London',                          'britain',   1),
    ('Germany Wozzeck Again',           'germany',   3),
    ('Paris Le Train Bleu',             'france',    2),
    ('Paris Le Train Bleu again',       'france',    3),
    ('Tokeo',                           'japan',     1),
    ('Bombay',                          'india',     1),
    ('Havana',                          'cuba',      1),
    ('Rio De Janeiro',                  'brazil',    1),
    ('Mexico City',                     'mexico',    1),
    ('Bali',                            'indonesia', 1),
]

# ── relationships: rId → media filename ──────────────────────────────────────
rels = {}
rels_tree = ET.parse(os.path.join(BASE, '_rels', 'document.xml.rels'))
for rel in rels_tree.getroot():
    rid = rel.get('Id')
    target = rel.get('Target', '')
    if 'image' in rel.get('Type', ''):
        rels[rid] = target          # e.g. "media/image5.jpeg"

# ── helpers ───────────────────────────────────────────────────────────────────
def get_images_in_elem(elem):
    images = []
    for blip in elem.iter(f'{{{A}}}blip'):
        rid = blip.get(f'{{{R}}}embed')
        if rid and rid in rels:
            images.append(rels[rid])
    for imgdata in elem.iter(f'{{{VML}}}imagedata'):
        rid = imgdata.get(f'{{{R}}}id')
        if rid and rid in rels:
            images.append(rels[rid])
    return images

def render_run(run):
    """Return inline HTML for a <w:r> run (bold / italic aware)."""
    rpr = run.find(f'{{{W}}}rPr')
    bold   = rpr is not None and rpr.find(f'{{{W}}}b')  is not None
    italic = rpr is not None and rpr.find(f'{{{W}}}i')  is not None
    text = ''.join(t.text or '' for t in run.findall(f'{{{W}}}t'))
    if not text:
        return ''
    text = (text.replace('&', '&amp;')
                .replace('<', '&lt;')
                .replace('>', '&gt;'))
    if bold and italic:
        text = f'<strong><em>{text}</em></strong>'
    elif bold:
        text = f'<strong>{text}</strong>'
    elif italic:
        text = f'<em>{text}</em>'
    return text

def para_to_html(para):
    """Convert a <w:p> to an HTML string (returns '' for blank paragraphs)."""
    parts = []
    for child in para:
        if child.tag == f'{{{W}}}r':
            parts.append(render_run(child))
        elif child.tag == f'{{{W}}}hyperlink':
            for run in child.findall(f'{{{W}}}r'):
                parts.append(render_run(run))
    text = ''.join(parts).strip()
    return text

# ── parse document ────────────────────────────────────────────────────────────
doc_tree = ET.parse(os.path.join(BASE, 'document.xml'))
body = doc_tree.getroot().find(f'.//{{{W}}}body')

# Collect vignette data in order
vignettes = []   # list of dicts: label, country, n, paragraphs, images
current = None
label_to_entry = {}

for (lbl, country, n) in VIGNETTE_MAP:
    label_to_entry[lbl] = None   # will be filled

for elem in body:
    tag = elem.tag

    if tag == f'{{{W}}}p':
        raw = ''.join(t.text or '' for t in elem.iter(f'{{{W}}}t')).strip()

        # Detect VIGNETTE_START
        m = re.match(r'VIGNETTE_START\s*[:\-]?\s*(.+)', raw, re.IGNORECASE)
        if m:
            lbl = m.group(1).strip()
            # Find mapping entry (fuzzy: strip case)
            matched = None
            for (doc_lbl, country, n) in VIGNETTE_MAP:
                if doc_lbl.lower() == lbl.lower():
                    matched = (doc_lbl, country, n)
                    break
            if matched:
                current = {'label': matched[0], 'country': matched[1],
                           'n': matched[2], 'paragraphs': [], 'images': []}
                vignettes.append(current)
            else:
                print(f'WARNING: no mapping for label "{lbl}"')
                current = None
            continue

        if re.match(r'VIGNETTE_END', raw, re.IGNORECASE):
            current = None
            continue

        if current is not None:
            html = para_to_html(elem)
            if html:
                current['paragraphs'].append(html)
            imgs = get_images_in_elem(elem)
            current['images'].extend(imgs)

    elif tag == f'{{{W}}}tbl':
        if current is not None:
            for cell_para in elem.iter(f'{{{W}}}p'):
                html = para_to_html(cell_para)
                if html:
                    current['paragraphs'].append(html)
                imgs = get_images_in_elem(cell_para)
                current['images'].extend(imgs)

print(f'Parsed {len(vignettes)} vignettes.\n')

# ── build per-country counts for manifest ─────────────────────────────────────
country_counts = {}
for v in vignettes:
    c = v['country']
    country_counts[c] = max(country_counts.get(c, 0), v['n'])

# ── write vignette HTML files & copy images ───────────────────────────────────
seen_media = {}   # media path → first output path (for dedup tracking only)

for v in vignettes:
    country = v['country']
    n       = v['n']
    slug    = f'{country}_{n}'

    # --- copy + rename images ---
    img_tags = []
    for pic_idx, media_rel in enumerate(v['images'], start=1):
        src_name = os.path.basename(media_rel)          # e.g. image5.jpeg
        src_path = os.path.join(MEDIA, src_name)
        dst_name = f'{slug}_pic{pic_idx}.jpg'
        dst_path = os.path.join(IMG_DIR, dst_name)
        if os.path.exists(src_path):
            shutil.copy2(src_path, dst_path)
        else:
            print(f'  MISSING: {src_path}')
        img_tags.append(f'    <img src="../images/{dst_name}" alt="" loading="lazy">')

    # --- build HTML fragment ---
    paras_html = '\n'.join(f'    <p>{p}</p>' for p in v['paragraphs'])
    imgs_html  = '\n'.join(img_tags)

    html = f'''<div class="vignette" data-slug="{slug}">
  <div class="vignette-images">
{imgs_html}
  </div>
  <div class="vignette-text">
{paras_html}
  </div>
</div>
'''
    out_path = os.path.join(VIG_DIR, f'{slug}.html')
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'  Wrote {slug}.html  ({len(v["images"])} images, {len(v["paragraphs"])} paragraphs)')

# ── blank shells for countries with no DOCX content ──────────────────────────
ALL_COUNTRIES = [
    'us','cuba','caribbean','argentina','brazil','mexico','colombia',
    'france','germany','ballet','britain','ireland','spain','italy',
    'portugal','russia','hungary','czechia','poland','balkans',
    'japan','china','india','indonesia','philippines','korea',
    'egypt','africa','middleeast','australia'
]
shells_written = []
for c in ALL_COUNTRIES:
    if c not in country_counts:
        path = os.path.join(VIG_DIR, f'{c}_1.html')
        html = f'''<div class="vignette vignette--placeholder" data-slug="{c}_1">
  <div class="vignette-text">
    <p class="vignette-coming-soon">Vignette coming soon.</p>
  </div>
</div>
'''
        with open(path, 'w', encoding='utf-8') as f:
            f.write(html)
        shells_written.append(f'{c}_1.html')
        country_counts[c] = 1

print(f'\nWrote {len(shells_written)} placeholder shells:')
for s in shells_written:
    print(f'  {s}')

# ── write manifest JSON ───────────────────────────────────────────────────────
manifest = {c: country_counts[c] for c in ALL_COUNTRIES if c in country_counts}
manifest_path = os.path.join(VIG_DIR, 'manifest.json')
with open(manifest_path, 'w', encoding='utf-8') as f:
    json.dump(manifest, f, indent=2)
print(f'\nWrote vignettes/manifest.json')
print(json.dumps(manifest, indent=2))
