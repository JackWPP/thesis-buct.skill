#!/usr/bin/env python3
"""Inject Pandoc OMML into styled DOCX. Handles block formulas (FRM_) and inline formulas (INL_)."""
import re, os, json, zipfile, subprocess
from lxml import etree as ET

WORKSPACE = os.path.dirname(os.path.abspath(__file__))
PANDOC = 'C:/Program Files/Pandoc/pandoc'
NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS_M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

for pfx, uri in [('w', NS_W), ('m', NS_M)]:
    ET.register_namespace(pfx, uri)

def w_tag(name): return f'{{{NS_W}}}{name}'

def paragraph_text(p):
    return ''.join(t.text or '' for t in p.iter(w_tag('t')))

def pandoc_omml(latex):
    """Run Pandoc on a single LaTeX formula via stdin."""
    docx_file = os.path.join(WORKSPACE, '_eq.docx')
    try:
        r = subprocess.run([PANDOC, '--from', 'markdown+tex_math_dollars', '--to', 'docx', '-o', docx_file],
                          input=f'$$\n{latex}\n$$\n', capture_output=True, text=True, timeout=15)
        if r.returncode != 0:
            return None, None
        with zipfile.ZipFile(docx_file, 'r') as z:
            tree = ET.fromstring(z.read('word/document.xml'))
        body = tree.find(w_tag('body'))
        omml_p, omml_elem = None, None
        for p in body.iter(w_tag('p')):
            if p.find(f'{{{NS_M}}}oMathPara') is not None or p.find(f'{{{NS_M}}}oMath') is not None:
                omml_p = p
                omml_elem = p.find(f'{{{NS_M}}}oMathPara') or p.find(f'{{{NS_M}}}oMath')
                break
        return omml_p, omml_elem
    except Exception as e:
        return None, None
    finally:
        try: os.unlink(docx_file)
        except: pass

def inject(docx_path, formulas_map, output_path):
    with zipfile.ZipFile(docx_path, 'r') as zin:
        data = {n: zin.read(n) for n in zin.namelist()}
    tree = ET.fromstring(data['word/document.xml'])
    body = tree.find(w_tag('body'))

    b_ok, b_fail, i_ok, i_fail = 0, 0, 0, 0
    cleaned = 0

    # ── Process EVERY run in the entire document in one pass ──
    all_runs = list(body.iter(w_tag('r')))

    for r in all_runs:
        # Get the text content of this run
        ts = list(r.iter(w_tag('t')))
        if not ts: continue
        t = ts[0]
        if not t.text: continue

        # ── Inline formula (INL_ marker) ──
        im = re.search(r'\[\[(INL_\d+)\]\]', t.text or '')
        if im:
            marker = im.group(1)
            info = formulas_map.get(marker)
            if not info or not info['latex'].strip():
                t.text = (t.text or '').replace(f'[[{marker}]]', '')
                cleaned += 1; continue

            latex = info['latex'].strip()
            pd_p, om = pandoc_omml(latex)
            if pd_p is None or om is None:
                i_fail += 1
                t.text = (t.text or '').replace(f'[[{marker}]]', '')
                continue

            # BEFORE modifications: find marker run and fallback runs in parent
            parent_el = r.getparent()
            if parent_el is None:
                i_fail += 1; continue

            siblings = list(parent_el)
            marker_pos = -1
            for si, sib in enumerate(siblings):
                if sib is r or sib == r:
                    marker_pos = si; break
            if marker_pos < 0:
                i_fail += 1; continue

            # Identify fallback runs: consecutive runs after marker with color=888888
            fallback_runs = []
            for si in range(marker_pos + 1, len(siblings)):
                sib = siblings[si]
                if sib.tag == w_tag('r'):
                    rPr = sib.find(w_tag('rPr'))
                    if rPr is not None:
                        c = rPr.find(w_tag('color'))
                        if c is not None and c.get(w_tag('val')) == '888888':
                            fallback_runs.append(sib)
                            continue
                break  # Stop at first non-fallback

            # Insert OMML before the marker run
            for child in list(pd_p):
                if child.tag == w_tag('pPr'): continue
                for sub in list(child):
                    parent_el.insert(marker_pos, sub)
                    marker_pos += 1

            # Remove marker + fallback runs
            try: parent_el.remove(r)
            except: pass
            for fb in fallback_runs:
                try: parent_el.remove(fb)
                except: pass
            i_ok += 1
            continue  # Skip block check for this run (already processed)

        # ── Block formula (FRM_ marker) ──
        fm = re.search(r'\[\[FRM_(\d+_\w+)\]\]', t.text or '')
        if fm:
            marker = f'FRM_{fm.group(1)}'
            info = formulas_map.get(marker)
            if info and info['latex'].strip() and info.get('eqLabel'):
                latex = info['latex'].strip()
                elabel = info['eqLabel']
                pd_p, om = pandoc_omml(latex)
                if pd_p is not None and om is not None:
                    try:
                        p = get_paragraph_ancestor(r)
                        if p is not None:
                            # Build new paragraph with OMML + equation number
                            new_p = ET.Element(w_tag('p'))
                            pPr = ET.SubElement(new_p, w_tag('pPr'))
                            ET.SubElement(pPr, w_tag('jc')).set(w_tag('val'), 'center')
                            sp = ET.SubElement(pPr, w_tag('spacing'))
                            for attr, val in [('line','440'),('lineRule','atLeast'),('before','200'),('after','200')]:
                                sp.set(w_tag(attr), val)
                            tabs = ET.SubElement(pPr, w_tag('tabs'))
                            for pos, al in [(4422, 'center'), (8844, 'right')]:
                                tab = ET.SubElement(tabs, w_tag('tab'))
                                tab.set(w_tag('val'), al); tab.set(w_tag('pos'), str(pos))
                            tr = ET.SubElement(new_p, w_tag('r'))
                            ET.SubElement(tr, w_tag('t')).text = '\t'
                            for child in list(pd_p):
                                if child.tag == w_tag('pPr'): continue
                                new_p.append(child)
                            if elabel:
                                nr = ET.SubElement(new_p, w_tag('r'))
                                nrPr = ET.SubElement(nr, w_tag('rPr'))
                                ET.SubElement(nrPr, w_tag('rFonts')).set(w_tag('ascii'), 'Times New Roman')
                                ET.SubElement(nr, w_tag('t')).text = f'\t（{elabel}）'
                            p.getparent().insert(list(p.getparent()).index(p), new_p)
                            p.getparent().remove(p)
                            b_ok += 1
                    except Exception:
                        b_fail += 1
                else:
                    b_fail += 1
            else:
                t.text = (t.text or '').replace(f'[[{marker}]]', '')
                cleaned += 1

    # ── Final cleanup: fix any remaining grey fallback text ──
    for r in body.iter(w_tag('r')):
        rPr = r.find(w_tag('rPr'))
        if rPr is not None:
            c = rPr.find(w_tag('color'))
            if c is not None and c.get(w_tag('val')) == '888888':
                c.set(w_tag('val'), 'auto')
    for r in body.iter(w_tag('r')):
        for t in r.iter(w_tag('t')):
            if t.text and '[[' in (t.text or ''):
                t.text = re.sub(r'\[\[(?:FRM_|INL_)\w+\]\]', '', t.text or '')
                cleaned += 1

    print(f'  Block: {b_ok} OK, {b_fail} failed')
    print(f'  Inline: {i_ok} OK, {i_fail} failed')
    print(f'  Cleaned: {cleaned} markers')
    data['word/document.xml'] = ET.tostring(tree, encoding='UTF-8', xml_declaration=True)
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, content in data.items():
            zout.writestr(name, content)
    print(f'  Output: {output_path} ({os.path.getsize(output_path)/1024/1024:.1f} MB)')

def get_paragraph_ancestor(el):
    """Find the nearest w:p ancestor of an element."""
    while el is not None:
        if el.tag == w_tag('p'):
            return el
        el = el.getparent()
    return None
    with zipfile.ZipFile(docx_path, 'r') as zin:
        data = {n: zin.read(n) for n in zin.namelist()}
    tree = ET.fromstring(data['word/document.xml'])
    body = tree.find(w_tag('body'))

    b_ok, b_fail, i_ok, i_fail = 0, 0, 0, 0

    for p in list(body.iter(w_tag('p'))):
        text = paragraph_text(p)

        # ── Block formula (FRM_ marker) ──
        bm = re.search(r'\[\[FRM_(\d+_\w+)\]\]', text)
        if bm:
            marker = f'FRM_{bm.group(1)}'
            info = formulas_map.get(marker)
            if not info: continue
            latex = info['latex'].strip()
            elabel = info['eqLabel']
            if not latex: continue

            pd_p, om = pandoc_omml(latex)
            if pd_p is None or om is None:
                b_fail += 1; print(f'  Block FAILED: {elabel or marker}'); continue

            # Build paragraph with CENTER + RIGHT tab stops
            new_p = ET.Element(w_tag('p'))
            pPr = ET.SubElement(new_p, w_tag('pPr'))
            ET.SubElement(pPr, w_tag('jc')).set(w_tag('val'), 'center')
            sp = ET.SubElement(pPr, w_tag('spacing'))
            for attr, val in [('line','400'),('lineRule','exact'),('before','200'),('after','200')]:
                sp.set(w_tag(attr), val)
            tabs = ET.SubElement(pPr, w_tag('tabs'))
            for pos, al in [(int(8844/2), 'center'), (8844, 'right')]:
                tab = ET.SubElement(tabs, w_tag('tab'))
                tab.set(w_tag('val'), al); tab.set(w_tag('pos'), str(pos))

            # Leading tab (centers formula)
            tr = ET.SubElement(new_p, w_tag('r'))
            ET.SubElement(tr, w_tag('t')).text = '\t'

            # Copy OMML from Pandoc output
            for child in list(pd_p):
                if child.tag == w_tag('pPr'): continue
                new_p.append(child)

            # Equation number: \t + （3-2） (RIGHT tab)
            if elabel:
                nr = ET.SubElement(new_p, w_tag('r'))
                nrPr = ET.SubElement(nr, w_tag('rPr'))
                nrfs = ET.SubElement(nrPr, w_tag('rFonts'))
                nrfs.set(w_tag('ascii'), 'Times New Roman')
                nrfs.set(w_tag('hAnsi'), 'Times New Roman')
                ET.SubElement(nr, w_tag('t')).text = f'\t（{elabel}）'

            p.getparent().insert(list(p.getparent()).index(p), new_p)
            p.getparent().remove(p)
            b_ok += 1
            continue

        # ── Inline formula (INL_ marker) ──
        im = re.search(r'\[\[(INL_\d+)\]\]', text)
        if im:
            marker = im.group(1)
            info = formulas_map.get(marker)
            if not info: continue
            latex = info['latex'].strip()
            if not latex: continue

            # Find the marker run (tiny, white) within the paragraph
            all_runs = list(p.iter(w_tag('r')))
            marker_ri = -1
            for ri, r in enumerate(all_runs):
                for t in r.iter(w_tag('t')):
                    if t.text and f'[[{marker}]]' in (t.text or ''):
                        marker_ri = ri
                        break
                if marker_ri >= 0: break

            if marker_ri < 0: continue

            # Find the fallback run (grey italic LaTeX source, immediately after marker)
            fallback_ri = -1
            for ri in range(marker_ri + 1, min(marker_ri + 3, len(all_runs))):
                r = all_runs[ri]
                rPr = r.find(w_tag('rPr'))
                has_color = False
                if rPr is not None:
                    color_el = rPr.find(w_tag('color'))
                    if color_el is not None:
                        has_color = color_el.get(w_tag('val'), '') == '888888'
                if has_color:
                    fallback_ri = ri
                    break

            # Run Pandoc to generate OMML
            pd_p, om = pandoc_omml(latex)
            if pd_p is not None and om is not None:
                # Insert OMML before the marker run
                ref_child = all_runs[marker_ri]
                ref_parent = ref_child.getparent()
                ref_idx = list(ref_parent).index(ref_child)
                for child in list(pd_p):
                    if child.tag == w_tag('pPr'): continue
                    for sub in list(child):
                        ref_parent.insert(ref_idx, sub)
                        ref_idx += 1
                # Remove marker run
                p.remove(all_runs[marker_ri])
                # Remove fallback run
                if fallback_ri >= 0:
                    fallback = all_runs[fallback_ri]
                    try: p.remove(fallback)
                    except: pass
                i_ok += 1
            else:
                # OMML failed — keep grey fallback text, just remove visible marker
                i_fail += 1
                p.remove(all_runs[marker_ri])
            continue

    # ── Final cleanup: remove all remaining markers and fix grey fallback text ──
    cleaned = 0
    for p in body.iter(w_tag('p')):
        for r in list(p.iter(w_tag('r'))):
            for t in list(r.iter(w_tag('t'))):
                if t.text and ('[[' in (t.text or '')):
                    if re.search(r'\[\[(?:FRM_|INL_)\w+\]\]', t.text):
                        t.text = re.sub(r'\[\[(?:FRM_|INL_)\w+\]\]', '', t.text or '')
                        cleaned += 1
                        if not t.text.strip():
                            try: p.remove(r)
                            except: pass
                # Fix grey fallback: change color to auto (black) for readability
                rPr = r.find(w_tag('rPr'))
                if rPr is not None:
                    color_el = rPr.find(w_tag('color'))
                    if color_el is not None and color_el.get(w_tag('val')) == '888888':
                        color_el.set(w_tag('val'), 'auto')

    print(f'  Block: {b_ok} OK, {b_fail} failed')
    print(f'  Inline: {i_ok} OK, {i_fail} failed')
    print(f'  Cleanup: {cleaned} markers removed, grey text fixed')

    data['word/document.xml'] = ET.tostring(tree, encoding='UTF-8', xml_declaration=True)
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, content in data.items():
            zout.writestr(name, content)
    print(f'  Output: {output_path} ({os.path.getsize(output_path)/1024/1024:.1f} MB)')

if __name__ == '__main__':
    styled = os.path.join(WORKSPACE, 'output', 'thesis_styled.docx')
    fmap = os.path.join(WORKSPACE, 'formulas_map.json')
    final = os.path.join(WORKSPACE, 'output', 'thesis_omml.docx')
    if not os.path.exists(styled):
        print(f'ERROR: Run build_thesis.js first')
        exit(1)
    formulas = json.load(open(fmap, 'r', encoding='utf-8'))
    print(f'Injecting {len(formulas)} formulas...')
    inject(styled, formulas, final)
    print('Complete.')
