#!/usr/bin/env python3
"""
Replace w:hyperlink with proper Word cross-reference fields (REF) that link to numbered list items.
"""
import zipfile, re, os, io, copy, shutil
from lxml import etree as ET

WORKSPACE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(WORKSPACE, 'output', 'thesis_omml.docx')
DST = os.path.join(WORKSPACE, 'output', 'thesis_crossref.docx')

NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
for pfx, uri in [('w', NS_W)]: ET.register_namespace(pfx, uri)

def w_tag(n): return f'{{{NS_W}}}{n}'

# Load DOCX
with zipfile.ZipFile(SRC, 'r') as zin:
    data = {n: zin.read(n) for n in zin.namelist()}

tree = ET.fromstring(data['word/document.xml'])
body = tree.find(w_tag('body'))

# ── Step 1: Add numbering definition to numbering.xml ──
# If numbering.xml doesn't exist, create it
num_xml = data.get('word/numbering.xml', b'')
if not num_xml:
    num_xml = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
             xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
</w:numbering>'''

num_tree = ET.fromstring(num_xml)

# Check if our abstract numbering already exists
existing = num_tree.findall(f'.//{{{NS_W}}}abstractNum')
next_abs_id = max([int(a.get(w_tag('abstractNumId'), '0')) for a in existing] + [0]) + 1

# Create abstract numbering for references
abs_num = ET.SubElement(num_tree, w_tag('abstractNum'))
abs_num.set(w_tag('abstractNumId'), str(next_abs_id))

# Multi-level type (but we only use level 0)
multi = ET.SubElement(abs_num, w_tag('multiLevelType'))
multi.set(w_tag('val'), 'hybridMultilevel')

# Level 0: [1], [2], etc. with hanging indent
for lvl_id, (fmt, start) in enumerate([('decimal', 1)]):
    lvl = ET.SubElement(abs_num, w_tag('lvl'))
    lvl.set(w_tag('ilvl'), str(lvl_id))
    start_el = ET.SubElement(lvl, w_tag('start'))
    start_el.set(w_tag('val'), str(start))
    numFmt = ET.SubElement(lvl, w_tag('numFmt'))
    numFmt.set(w_tag('val'), fmt)
    lvlText = ET.SubElement(lvl, w_tag('lvlText'))
    lvlText.set(w_tag('val'), '[%1]')
    lvlJc = ET.SubElement(lvl, w_tag('lvlJc'))
    lvlJc.set(w_tag('val'), 'left')
    # Paragraph properties for this level
    pPr = ET.SubElement(lvl, w_tag('pPr'))
    ind = ET.SubElement(pPr, w_tag('ind'))
    ind.set(w_tag('left'), '480')
    ind.set(w_tag('hanging'), '480')

# Create num instance linking to abstract num
next_num_id = max([int(n.get(w_tag('numId'), '0')) for n in num_tree.findall(f'.//{{{NS_W}}}num')] + [0]) + 1
num_inst = ET.SubElement(num_tree, w_tag('num'))
num_inst.set(w_tag('numId'), str(next_num_id))
ab_ref = ET.SubElement(num_inst, w_tag('abstractNumId'))
ab_ref.set(w_tag('val'), str(next_abs_id))

data['word/numbering.xml'] = ET.tostring(num_tree, encoding='UTF-8', xml_declaration=True)

# ── Step 2: Replace bookmark-only approach with numbered items ──
bm_counter = 0
ref_count = 0

for p in list(body.iter(w_tag('p'))):
    # Find reference paragraphs (contain bookmarks + cite numbers)
    bms = p.findall(w_tag('bookmarkStart'))
    if not bms: continue

    for bm in bms:
        bm_name = bm.get(w_tag('name'), '')
        if not bm_name.startswith('ref_R'): continue

        # This is a reference paragraph — add w:numPr for auto-numbering
        pPr = p.find(w_tag('pPr'))
        if pPr is None:
            pPr = ET.Element(w_tag('pPr'))
            p.insert(0, pPr)

        # Remove existing manual number runs ([n] text)
        for r in list(p.iter(w_tag('r'))):
            for t in r.iter(w_tag('t')):
                if t.text and re.match(r'^\[\d+\]\s*$', (t.text or '')):
                    # Remove this run (auto-numbering replaces it)
                    p.remove(r)
                    break

        # Add numPr if not present
        numPr = pPr.find(w_tag('numPr'))
        if numPr is None:
            numPr = ET.Element(w_tag('numPr'))
            pPr.insert(0, numPr)
        # Clear existing numPr children
        for c in list(numPr):
            numPr.remove(c)
        ilvl = ET.SubElement(numPr, w_tag('ilvl'))
        ilvl.set(w_tag('val'), '0')
        numId = ET.SubElement(numPr, w_tag('numId'))
        numId.set(w_tag('val'), str(next_num_id))

        ref_count += 1

print(f'  References numbered: {ref_count}')

# ── Step 3: Replace w:hyperlink with REF field codes ──
cite_count = 0
for p in body.iter(w_tag('p')):
    for hl in list(p.iter(w_tag('hyperlink'))):
        anchor = hl.get(w_tag('anchor'), '')
        if not anchor.startswith('ref_R'): continue

        # Find the index of this hyperlink in its parent
        parent = hl.getparent()
        idx = list(parent).index(hl)

        # Create REF field code
        # BEGIN
        begin_r = ET.Element(w_tag('r'))
        begin_fld = ET.SubElement(begin_r, w_tag('fldChar'))
        begin_fld.set(w_tag('fldCharType'), 'begin')

        # INSTRUCTION
        instr_r = ET.Element(w_tag('r'))
        instr_t = ET.SubElement(instr_r, w_tag('instrText'))
        instr_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        instr_t.text = f' REF {anchor} \\n \\h '

        # SEPARATE
        sep_r = ET.Element(w_tag('r'))
        sep_fld = ET.SubElement(sep_r, w_tag('fldChar'))
        sep_fld.set(w_tag('fldCharType'), 'separate')

        # Content: copy the superscript text from the hyperlink
        content_r = ET.Element(w_tag('r'))
        rPr = ET.SubElement(content_r, w_tag('rPr'))
        rFonts = ET.SubElement(rPr, w_tag('rFonts'))
        rFonts.set(w_tag('ascii'), 'Times New Roman')
        rFonts.set(w_tag('hAnsi'), 'Times New Roman')
        va = ET.SubElement(rPr, w_tag('vertAlign'))
        va.set(w_tag('val'), 'superscript')
        sz = ET.SubElement(rPr, w_tag('sz'))
        sz.set(w_tag('val'), '24')
        # Get the cite number from the hyperlink
        cite_text = ''
        for t in hl.iter(w_tag('t')):
            cite_text = (t.text or '')
            break
        content_t = ET.SubElement(content_r, w_tag('t'))
        content_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        content_t.text = cite_text

        # END
        end_r = ET.Element(w_tag('r'))
        end_fld = ET.SubElement(end_r, w_tag('fldChar'))
        end_fld.set(w_tag('fldCharType'), 'end')

        # Remove old hyperlink
        parent.remove(hl)

        # Insert field code elements at the same position
        for el in [begin_r, instr_r, sep_r, content_r, end_r]:
            parent.insert(idx, el)
            idx += 1

        cite_count += 1

print(f'  Citations converted to REF fields: {cite_count}')

# ── Step 4: Also update [Content_Types].xml if needed ──
# The numbering.xml reference must exist in [Content_Types].xml
ct_xml = data['[Content_Types].xml']
ct_tree = ET.fromstring(ct_xml)
ct_ns = 'http://schemas.openxmlformats.org/package/2006/content-types'
has_num = any(o.get('PartName') == '/word/numbering.xml' for o in ct_tree)
if not has_num:
    ov = ET.SubElement(ct_tree, f'{{{ct_ns}}}Override')
    ov.set('PartName', '/word/numbering.xml')
    ov.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml')
    data['[Content_Types].xml'] = ET.tostring(ct_tree, encoding='UTF-8', xml_declaration=True)

# ── Step 5: Write output ──
data['word/document.xml'] = ET.tostring(tree, encoding='UTF-8', xml_declaration=True)

with zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED) as zout:
    for name, content in data.items():
        zout.writestr(name, content)

print(f'  Output: {DST}')
