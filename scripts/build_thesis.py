#!/usr/bin/env python3
"""
BUCT Thesis DOCX Builder v4 — Pandoc + Post-processing Pipeline
Requires: pandoc, lxml, Pillow
"""
import re, os, json, zipfile, shutil, subprocess, io, copy, struct
from lxml import etree as ET
from PIL import Image

WORKSPACE = os.path.dirname(os.path.abspath(__file__))
PANDOC = 'C:/Program Files/Pandoc/pandoc'

NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
NS_WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
NS_R2 = 'http://schemas.openxmlformats.org/package/2006/relationships'

for prefix, uri in [('w', NS_W), ('m', NS_M), ('wp', NS_WP), ('a', NS_A),
                     ('pic', NS_PIC), ('r', NS_R2)]:
    ET.register_namespace(prefix, uri)

# ═══════════════════════════════════════════════════════════════════════════════
# Helpers to build OOXML elements
# ═══════════════════════════════════════════════════════════════════════════════

def w_tag(name): return f'{{{NS_W}}}{name}'

def make_run(text, bold=False, italics=False, superscript=False, size_pt=24):
    r = ET.Element(w_tag('r'))
    rPr = ET.SubElement(r, w_tag('rPr'))
    rFonts = ET.SubElement(rPr, w_tag('rFonts'))
    rFonts.set(w_tag('eastAsia'), '宋体' if any('一' <= c <= '鿿' for c in (text or '')) else 'Times New Roman')
    for fn in ['ascii', 'hAnsi', 'cs']:
        rFonts.set(w_tag(fn), 'Times New Roman')
    for sz_tag in [w_tag('sz'), w_tag('szCs')]:
        ET.SubElement(rPr, sz_tag).set(w_tag('val'), str(size_pt))
    if bold: ET.SubElement(rPr, w_tag('b'))
    if italics: ET.SubElement(rPr, w_tag('i'))
    if superscript:
        va = ET.SubElement(rPr, w_tag('vertAlign'))
        va.set(w_tag('val'), 'superscript')
    t = ET.SubElement(r, w_tag('t'))
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = text
    return r

def make_para(children=None, align='both', first_indent=480, before=0, after=0, spacing_line=400):
    p = ET.Element(w_tag('p'))
    pPr = ET.SubElement(p, w_tag('pPr'))
    jc = ET.SubElement(pPr, w_tag('jc'))
    jc.set(w_tag('val'), {'left':'left','center':'center','right':'right','both':'both'}[align])
    sp = ET.SubElement(pPr, w_tag('spacing'))
    sp.set(w_tag('line'), str(spacing_line)); sp.set(w_tag('lineRule'), 'exact')
    sp.set(w_tag('before'), str(before)); sp.set(w_tag('after'), str(after))
    if first_indent:
        ET.SubElement(pPr, w_tag('ind')).set(w_tag('firstLine'), str(first_indent))
    if children:
        for c in children:
            p.append(c)
    return p

def make_hyperlink(anchor, runs, superscript=False):
    hl = ET.Element(w_tag('hyperlink'))
    hl.set(w_tag('anchor'), anchor)
    hl.set(w_tag('history'), '1')
    for r in runs:
        if superscript:
            rPr = r.find(w_tag('rPr'))
            if rPr is None:
                rPr = ET.Element(w_tag('rPr'))
                r.insert(0, rPr)
            va = ET.SubElement(rPr, w_tag('vertAlign'))
            va.set(w_tag('val'), 'superscript')
        hl.append(r)
    return hl

def make_image_drawing(rId, name, width_emu, height_emu):
    """Create a w:drawing element for an inline image."""
    wp_extent_cx = width_emu
    wp_extent_cy = height_emu

    drawing = ET.Element(w_tag('drawing'))
    inline = ET.SubElement(drawing, f'{{{NS_WP}}}inline')
    inline.set('distT', '0'); inline.set('distB', '0'); inline.set('distL', '0'); inline.set('distR', '0')

    extent = ET.SubElement(inline, f'{{{NS_WP}}}extent')
    extent.set('cx', str(wp_extent_cx)); extent.set('cy', str(wp_extent_cy))

    effectExtent = ET.SubElement(inline, f'{{{NS_WP}}}effectExtent')
    effectExtent.set('l', '0'); effectExtent.set('t', '0'); effectExtent.set('r', '0'); effectExtent.set('b', '0')

    docPr = ET.SubElement(inline, f'{{{NS_WP}}}docPr')
    docPr.set('id', '1'); docPr.set('name', name)

    cNvGraphicFramePr = ET.SubElement(inline, f'{{{NS_WP}}}cNvGraphicFramePr')

    graphic = ET.SubElement(inline, f'{{{NS_A}}}graphic')
    graphicData = ET.SubElement(graphic, f'{{{NS_A}}}graphicData')
    graphicData.set('uri', 'http://schemas.openxmlformats.org/drawingml/2006/picture')

    pic = ET.SubElement(graphicData, f'{{{NS_PIC}}}pic')
    nvPicPr = ET.SubElement(pic, f'{{{NS_PIC}}}nvPicPr')
    cNvPr = ET.SubElement(nvPicPr, f'{{{NS_PIC}}}cNvPr')
    cNvPr.set('id', '0'); cNvPr.set('name', name)
    cNvPicPr = ET.SubElement(nvPicPr, f'{{{NS_PIC}}}cNvPicPr')

    blipFill = ET.SubElement(pic, f'{{{NS_PIC}}}blipFill')
    blip = ET.SubElement(blipFill, f'{{{NS_A}}}blip')
    blip.set(f'{{{NS_R2}}}embed', rId)
    stretch = ET.SubElement(blipFill, f'{{{NS_A}}}stretch')
    ET.SubElement(stretch, f'{{{NS_A}}}fillRect')

    spPr = ET.SubElement(pic, f'{{{NS_PIC}}}spPr')
    xfrm = ET.SubElement(spPr, f'{{{NS_A}}}xfrm')
    off = ET.SubElement(xfrm, f'{{{NS_A}}}off'); off.set('x', '0'); off.set('y', '0')
    ext = ET.SubElement(xfrm, f'{{{NS_A}}}ext'); ext.set('cx', str(wp_extent_cx)); ext.set('cy', str(wp_extent_cy))
    prstGeom = ET.SubElement(spPr, f'{{{NS_A}}}prstGeom'); prstGeom.set('prst', 'rect')
    ET.SubElement(prstGeom, f'{{{NS_A}}}avLst')

    return drawing

def make_bookmark_start(name, bid):
    bm = ET.Element(w_tag('bookmarkStart'))
    bm.set(w_tag('id'), str(bid)); bm.set(w_tag('name'), name)
    return bm

def make_bookmark_end(bid):
    bm = ET.Element(w_tag('bookmarkEnd'))
    bm.set(w_tag('id'), str(bid))
    return bm

def has_marker(para, prefix):
    for t in para.iter(w_tag('t')):
        if t.text and t.text.startswith(prefix):
            return t.text
    return None

def replace_para(old_para, new_elements):
    parent = old_para.getparent()
    idx = list(parent).index(old_para)
    for i, el in enumerate(new_elements):
        parent.insert(idx + i, el)
    parent.remove(old_para)

# ═══════════════════════════════════════════════════════════════════════════════
# Step 1: Merge chapters
# ═══════════════════════════════════════════════════════════════════════════════

def merge_chapters():
    chapters_dir = os.path.join(WORKSPACE, 'chapters')
    files = sorted([f for f in os.listdir(chapters_dir) if f.endswith('.md')])
    combined = []
    for f in files:
        content = open(os.path.join(chapters_dir, f), 'r', encoding='utf-8').read()
        content = re.sub(r'<!--\s*FIGURE:\s*(\S+)\s*-->', r'\n\n<div>FIGMARKER_\1</div>\n\n', content)
        content = re.sub(r'<!--\s*TABLE:\s*(\S+)\s*-->', r'\n\n<div>TBLMARKER_\1</div>\n\n', content)
        content = re.sub(r'<!--\s*EQ:\s*(\S+)\s*-->', '', content)
        content = re.sub(r'<!--\s*(CHAPTER|TITLE|REFS_USED|FIGURES_USED|TABLES_USED):[^>]*-->', '', content)
        combined.append(content)
    merged = '\n\n'.join(combined)
    md_path = os.path.join(WORKSPACE, 'merged_thesis.md')
    with open(md_path, 'w', encoding='utf-8') as fh:
        fh.write(merged)
    print(f'Merged {len(files)} chapters, {len(merged)} chars')
    return md_path

# ═══════════════════════════════════════════════════════════════════════════════
# Step 2: Pandoc
# ═══════════════════════════════════════════════════════════════════════════════

def run_pandoc(md_path):
    output = os.path.join(WORKSPACE, 'output', 'pandoc_thesis.docx')
    os.makedirs(os.path.dirname(output), exist_ok=True)
    r = subprocess.run([PANDOC, md_path, '--from', 'markdown+tex_math_dollars+raw_html',
                        '--to', 'docx', '-o', output], capture_output=True, text=True)
    if r.returncode != 0:
        raise RuntimeError(f'Pandoc failed: {r.stderr}')
    print(f'Pandoc -> {output}')
    return output

# ═══════════════════════════════════════════════════════════════════════════════
# Step 3: Post-processing
# ═══════════════════════════════════════════════════════════════════════════════

def post_process(pandoc_path):
    registry = json.load(open(os.path.join(WORKSPACE, 'figure_table_registry.json'), 'r', encoding='utf-8'))
    refs = json.load(open(os.path.join(WORKSPACE, 'references.json'), 'r', encoding='utf-8'))
    ref_map = {r['id']: r for r in refs['refs']}

    with zipfile.ZipFile(pandoc_path, 'r') as zin:
        docx = {n: zin.read(n) for n in zin.namelist()}

    tree = ET.fromstring(docx['word/document.xml'])
    body = tree.find(w_tag('body'))

    # Track next IDs
    next_rel_id = 100  # Start high to avoid conflicts with Pandoc's IDs
    next_bm_id = 1000

    # ---- 3a: Process FIGURE markers (embed actual images) ----
    image_rels = {}  # rId -> image path
    media_files = {}  # filename -> data

    fig_paras = [(p, has_marker(p, 'FIGMARKER_')) for p in body.iter(w_tag('p'))]
    fig_paras = [(p, m.replace('FIGMARKER_', '')) for p, m in fig_paras if m]

    print(f'  Processing {len(fig_paras)} figures...')
    for para, key in fig_paras:
        fig = registry['figures'].get(key)
        if not fig:
            replace_para(para, [make_para([make_run(f'[Figure not found: {key}]')])])
            continue

        img_path = os.path.join(WORKSPACE, fig['file'])
        if not os.path.exists(img_path):
            replace_para(para, [make_para([make_run(f'[Image missing: {fig["file"]}]')])])
            continue

        # Read image and get dimensions
        img = Image.open(img_path)
        img_w_px, img_h_px = img.size

        # Content width in EMU for inline images
        width_ratio = fig.get('width_ratio', 0.7)
        content_w_emu = 8844 * 635  # 8844 DXA * 635 EMU/DXA
        img_w_emu = int(content_w_emu * width_ratio)
        img_h_emu = int(img_w_emu * (img_h_px / img_w_px))

        # Add image to media
        img_ext = os.path.splitext(img_path)[1].lower()
        media_name = f'image_fig_{key}{img_ext}'
        if media_name not in media_files:
            with open(img_path, 'rb') as fh:
                media_files[media_name] = fh.read()

        # Create relationship
        rId = f'rId_fig_{key}'
        image_rels[rId] = (media_name, f'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')

        # Build paragraphs: image + Chinese caption + English caption
        img_p = ET.Element(w_tag('p'))
        img_pPr = ET.SubElement(img_p, w_tag('pPr'))
        ET.SubElement(img_pPr, w_tag('jc')).set(w_tag('val'), 'center')
        img_sp = ET.SubElement(img_pPr, w_tag('spacing'))
        img_sp.set(w_tag('before'), '200')
        img_drawing = make_image_drawing(rId, f'fig_{key}', img_w_emu, img_h_emu)
        img_p.append(img_drawing)

        cn_p = make_para([
            make_run(f'{fig["label"]} ', size_pt=21),
            make_run(fig['caption_zh'], size_pt=21),
        ], align='center', first_indent=0, spacing_line=400)

        en_label = fig['label'].replace('图', '')
        en_p = make_para([
            make_run(f'Figure {en_label} {fig["caption_en"]}', size_pt=21),
        ], align='center', first_indent=0, spacing_line=400)

        replace_para(para, [img_p, cn_p, en_p])

    # ---- 3b: Process TABLE markers ----
    tbl_paras = [(p, has_marker(p, 'TBLMARKER_')) for p in body.iter(w_tag('p'))]
    tbl_paras = [(p, m.replace('TBLMARKER_', '')) for p, m in tbl_paras if m]

    print(f'  Processing {len(tbl_paras)} tables...')
    for para, key in tbl_paras:
        tab = registry['tables'].get(key)
        if not tab:
            replace_para(para, [make_para([make_run(f'[Table not found: {key}]')])])
            continue

        rows = tab.get('data', [])
        cols = tab.get('columns', [])
        if not rows:
            replace_para(para, [make_para([make_run(f'[Table data missing: {key}]')])])
            continue

        new_els = []
        # Captions
        new_els.append(make_para([make_run(f'{tab["label"]} ', size_pt=21), make_run(tab['caption_zh'], size_pt=21)],
                                  align='center', first_indent=0))
        el = tab['label'].replace('表', '')
        new_els.append(make_para([make_run(f'Table {el} {tab["caption_en"]}', size_pt=21)],
                                  align='center', first_indent=0))

        # Table element
        ncols = len(cols) if cols else (len(rows[0]) if rows else 1)
        cw = 8844 // ncols
        tbl = ET.Element(w_tag('tbl'))
        tblPr = ET.SubElement(tbl, w_tag('tblPr'))
        ET.SubElement(tblPr, w_tag('tblW')).set(w_tag('w'), '8844')
        ET.SubElement(tblPr, w_tag('tblW')).set(w_tag('type'), 'dxa')
        ET.SubElement(tblPr, w_tag('jc')).set(w_tag('val'), 'center')
        borders = ET.SubElement(tblPr, w_tag('tblBorders'))
        for bn in ['top','bottom','left','right','insideH','insideV']:
            b = ET.SubElement(borders, w_tag(bn))
            b.set(w_tag('val'), 'single'); b.set(w_tag('sz'), '6')
            b.set(w_tag('space'), '0'); b.set(w_tag('color'), '000000')
        tgrid = ET.SubElement(tbl, w_tag('tblGrid'))
        for ci in range(ncols):
            ET.SubElement(tgrid, w_tag('gridCol')).set(w_tag('w'), str(cw))

        # Header
        if cols:
            tr = ET.SubElement(tbl, w_tag('tr'))
            ET.SubElement(ET.SubElement(tr, w_tag('trPr')), w_tag('tblHeader'))
            for ci, ct in enumerate(cols):
                tc = ET.SubElement(tr, w_tag('tc'))
                tcPr = ET.SubElement(tc, w_tag('tcPr'))
                ET.SubElement(tcPr, w_tag('tcW')).set(w_tag('w'), str(cw))
                tc.append(make_para([make_run(str(ct), bold=True, size_pt=21)], align='center', first_indent=0))

        # Data rows
        for row in rows:
            tr = ET.SubElement(tbl, w_tag('tr'))
            for ci, ct in enumerate(row):
                tc = ET.SubElement(tr, w_tag('tc'))
                tcPr = ET.SubElement(tc, w_tag('tcPr'))
                ET.SubElement(tcPr, w_tag('tcW')).set(w_tag('w'), str(cw))
                tc.append(make_para([make_run(str(ct) if ct else '', size_pt=21)], align='center', first_indent=0))

        new_els.append(tbl)
        replace_para(para, new_els)

    # ---- 3c: Fix citation markers → cross-reference hyperlinks ----
    cite_pattern = re.compile(r'\[\^ref_(R\d+)\]')
    cite_count = 0
    for p in body.iter(w_tag('p')):
        for r in list(p):
            if r.tag != w_tag('r'):
                continue
            new_runs = []
            for child in list(r):
                if child.tag == w_tag('t') and child.text and cite_pattern.search(child.text):
                    text = child.text
                    last_end = 0
                    rPr = r.find(w_tag('rPr'))
                    for m in cite_pattern.finditer(text):
                        # Text before citation
                        if m.start() > last_end:
                            before = make_run(text[last_end:m.start()], size_pt=24)
                            if rPr is not None:
                                # Copy rPr from parent
                                pass
                            new_runs.append(before)
                        # Citation hyperlink
                        ref_id = f'ref_{m.group(1)}'
                        ref = ref_map.get(ref_id, {})
                        ck = ref.get('cite_key', '?')
                        cit_run = make_run(f'[{ck}]', size_pt=24)
                        hl = make_hyperlink(ref_id, [cit_run], superscript=True)
                        new_runs.append(hl)
                        cite_count += 1
                        last_end = m.end()
                    if last_end < len(text):
                        after = make_run(text[last_end:], size_pt=24)
                        new_runs.append(after)
                    # Replace children of this run
                    for c in list(r):
                        r.remove(c)
                    for nr in new_runs:
                        r.append(nr)
                    break  # Only process first t element with citations

    print(f'  Converted {cite_count} citations to hyperlinks')

    # ---- 3d: Build reference section ----
    # Remove any existing "References" heading from Pandoc
    ref_heading_found = False
    for p in list(body.iter(w_tag('p'))):
        for t in p.iter(w_tag('t')):
            if t.text and ('参考文献' in t.text or 'References' == t.text.strip()):
                # Remove this and all following elements (Pandoc's reference list)
                parent = p.getparent()
                idx = list(parent).index(p)
                to_remove = list(parent)[idx:]
                for el in to_remove:
                    parent.remove(el)
                ref_heading_found = True
                break
        if ref_heading_found:
            break

    # Add page break before references
    pb_para = ET.Element(w_tag('p'))
    pb_r = ET.SubElement(pb_para, w_tag('r'))
    ET.SubElement(pb_r, w_tag('br')).set(w_tag('type'), 'page')
    body.append(pb_para)

    # Reference title
    body.append(make_para([make_run('参考文献', size_pt=32)], align='center', first_indent=0, after=800))

    sorted_refs = sorted([r for r in refs['refs'] if r.get('cite_key', 0) > 0], key=lambda r: r['cite_key'])
    for ref in sorted_refs:
        formatted = ref.get('formatted', f'[{ref["cite_key"]}] {ref.get("title", "")}')
        bid = next_bm_id
        next_bm_id += 1
        ref_p = make_para(children=[], align='both', first_indent=0)
        ref_pPr = ref_p.find(w_tag('pPr'))
        ind = ref_pPr.find(w_tag('ind'))
        if ind is None:
            ind = ET.SubElement(ref_pPr, w_tag('ind'))
        ind.set(w_tag('left'), '480'); ind.set(w_tag('hanging'), '480')

        ref_p.append(make_bookmark_start(ref['id'], bid))
        ref_p.append(make_run(f'[{ref["cite_key"]}] ', size_pt=24))
        ref_p.append(make_run(formatted, size_pt=24))
        ref_p.append(make_bookmark_end(bid))
        body.append(ref_p)

    print(f'  Added {len(sorted_refs)} reference entries')

    # ---- 3e: Update relationships (document.xml.rels) ----
    if image_rels:
        rels_xml = docx.get('word/_rels/document.xml.rels', b'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
        rels_tree = ET.fromstring(rels_xml)
        for rId, (target, rtype) in image_rels.items():
            rel = ET.SubElement(rels_tree, '{http://schemas.openxmlformats.org/package/2006/relationships}Relationship')
            rel.set('Id', rId)
            rel.set('Type', rtype)
            rel.set('Target', f'media/{target}')
        docx['word/_rels/document.xml.rels'] = ET.tostring(rels_tree, encoding='UTF-8', xml_declaration=True)

    # Add media files
    for media_name, data in media_files.items():
        docx[f'word/media/{media_name}'] = data

    # Need to update [Content_Types].xml for image types
    ct_xml = docx['[Content_Types].xml']
    ct_tree = ET.fromstring(ct_xml)
    ct_ns = 'http://schemas.openxmlformats.org/package/2006/content-types'
    existing = set()
    for o in ct_tree:
        existing.add(o.get('PartName', ''))
    for media_name in media_files:
        ext = os.path.splitext(media_name)[1].lower()
        part_name = f'/word/media/{media_name}'
        if part_name not in existing:
            ctype = 'image/png' if ext == '.png' else 'image/jpeg'
            ov = ET.SubElement(ct_tree, f'{{{ct_ns}}}Override')
            ov.set('PartName', part_name)
            ov.set('ContentType', ctype)
            existing.add(part_name)
    docx['[Content_Types].xml'] = ET.tostring(ct_tree, encoding='UTF-8', xml_declaration=True)

    # ---- Write final DOCX ----
    output = os.path.join(WORKSPACE, 'output', 'thesis_final.docx')
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in docx.items():
            zout.writestr(name, data)

    size_mb = os.path.getsize(output) / (1024 * 1024)
    print(f'\nFinal DOCX -> {output} ({size_mb:.1f} MB)')
    return output

# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════

if __name__ == '__main__':
    print('=' * 60)
    print('BUCT Thesis Builder v4')
    print('=' * 60)
    md = merge_chapters()
    pandoc_out = run_pandoc(md)
    final = post_process(pandoc_out)
    print('Complete!')
