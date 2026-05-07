#!/usr/bin/env node
/**
 * build_thesis.js — BUCT 毕业论文 DOCX 构建 (v3: docx npm + formula markers)
 *
 * 完整控制格式（字体/边距/行距/页眉页脚）、图表、交叉引用。
 * 公式以 @@EQ_MARKER 占位，配套 formulas_map.json 供 Python 后处理注入 OMML。
 */
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, AlignmentType, BorderStyle, WidthType, VerticalAlign,
  HeadingLevel, PageBreak, LineRuleType, TabStopType, TabStopPosition,
  ShadingType, Header, Footer, PageNumber,
  Bookmark, InternalHyperlink,
} = require('docx');
const fs = require('fs');
const path = require('path');

const WORKSPACE = process.argv[3] || '.';
const OUTPUT    = process.argv[5] || './output/thesis_styled.docx';
const REGISTRY_PATH   = path.join(WORKSPACE, 'figure_table_registry.json');
const REFERENCES_PATH = path.join(WORKSPACE, 'references.json');
const CHAPTERS_DIR    = path.join(WORKSPACE, 'chapters');
const FORMULAS_MAP    = path.join(WORKSPACE, 'formulas_map.json');

const PAGE_WIDTH = 11906, PAGE_HEIGHT = 16838, MARGIN_TB = 1418, MARGIN_LR = 1531;
const CONTENT_WIDTH = PAGE_WIDTH - MARGIN_LR * 2;
const SZ_3 = 32, SZ_4 = 28, SZ_X4 = 24, SZ_5 = 21, SZ_X5 = 18;

let currentChapter = 0;
const formulaMap = {};  // "EQ_3_1" → { latex: "...", eqLabel: "3-1" }

// ─── Font helpers ───────────────────────────────────────────────────────────

function mixedTextRuns(text, size = SZ_X4, options = {}) {
  if (!text) return [];
  const runs = [];
  let cur = '', curCN = null;
  const flush = () => {
    if (!cur) return;
    runs.push(new TextRun({ text: cur, font: { name: curCN ? '宋体' : 'Times New Roman' }, size, ...options }));
    cur = '';
  };
  for (const ch of String(text)) {
    const cn = /[一-鿿　-〿＀-￯]/.test(ch);
    if (curCN !== null && cn !== curCN) flush();
    curCN = cn; cur += ch;
  }
  flush();
  return runs;
}

function headingRuns(text, size) {
  return String(text).split(/([a-zA-Z0-9.\-()/]+)/).filter(Boolean).map(p =>
    new TextRun({ text: p, font: { name: /^[a-zA-Z0-9.\-()/]+$/.test(p) ? 'Times New Roman' : '黑体' }, size })
  );
}

// ─── Paragraph factories ────────────────────────────────────────────────────

function chapterHeading(text) {
  const m = text.match(/^第(\d+)章/); if (m) currentChapter = parseInt(m[1]);
  return new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: 400, lineRule: 'atLeast', before: 0, after: 800 }, children: headingRuns(text, SZ_3) });
}
function sectionHeading(text) {
  return new Paragraph({ alignment: AlignmentType.LEFT, spacing: { line: 400, lineRule: 'atLeast', before: 0, after: 400 }, children: headingRuns(text, SZ_4) });
}
function subsectionHeading(text) {
  return new Paragraph({ alignment: AlignmentType.LEFT, spacing: { line: 400, lineRule: 'atLeast', before: 0, after: 400 }, children: headingRuns(text, SZ_X4) });
}
function subsubsectionHeading(text) {
  return new Paragraph({ alignment: AlignmentType.LEFT, spacing: { line: 400, lineRule: 'atLeast', before: 0, after: 0 }, children: headingRuns(text, SZ_X4) });
}

// ─── Formula — marker placeholder ───────────────────────────────────────────

function formulaMarker(latexStr, eqLabel) {
  // Use counter-based keys to avoid overwrites
  const mkBase = eqLabel ? eqLabel.replace(/-/g, '_') : 'auto';
  let key = `FRM_${mkBase}`;
  let counter = 1;
  while (formulaMap[key]) { key = `FRM_${mkBase}_${counter++}`; }
  formulaMap[key] = { latex: latexStr, eqLabel };

  const runs = mixedTextRuns(latexStr, SZ_X4);

  // Equation number — only for labeled formulas
  if (eqLabel) {
    runs.push(new TextRun({ text: `\t（${eqLabel}）`, font: { name: 'Times New Roman' }, size: SZ_X4 }));
    // Hidden marker ONLY if labeled (for OMML injection)
    runs.push(new TextRun({ text: `[[${key}]]`, font: { name: 'Times New Roman' }, size: 2, color: 'FFFFFF' }));
  }

  // Tab stops: CENTER for formula, RIGHT for number
  const tabStops = [
    { type: TabStopType.CENTER, position: CONTENT_WIDTH / 2 },
    { type: TabStopType.RIGHT, position: CONTENT_WIDTH },
  ];

  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: 400, lineRule: 'atLeast', before: 200, after: 200 },
    tabStops,
    children: [new TextRun({ text: '\t', font: { name: 'Times New Roman' }, size: SZ_X4 }), ...runs],
  });
}

// ─── Image insertion ────────────────────────────────────────────────────────

function insertFigure(figKey, registry) {
  const fig = registry.figures[figKey];
  if (!fig) { console.warn(`[WARN] Figure not found: ${figKey}`); return []; }
  const imgPath = path.join(WORKSPACE, fig.file);
  if (!fs.existsSync(imgPath)) {
    console.warn(`[WARN] Figure file not found: ${imgPath}`);
    return [new Paragraph({ children: [new TextRun({ text: `[图片缺失: ${fig.file}]`, color: 'FF0000' })] })];
  }
  const imgBuffer = fs.readFileSync(imgPath);
  const ext = path.extname(imgPath).slice(1).toLowerCase();
  const contentWidthPx = Math.round(CONTENT_WIDTH / 1440 * 96);
  const imgWidthPx = Math.round(contentWidthPx * (fig.width_ratio || 0.7));
  const imgHeightPx = Math.round(imgWidthPx * (fig.actual_ratio || 0.6));

  const pars = [];
  pars.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 0 }, children: [new ImageRun({ data: imgBuffer, transformation: { width: imgWidthPx, height: imgHeightPx }, type: ext === 'jpg' ? 'jpg' : 'png' })] }));
  pars.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: 400, lineRule: 'atLeast', before: 0, after: 0 }, children: [new TextRun({ text: `${fig.label} `, font: { name: '宋体' }, size: SZ_5 }), new TextRun({ text: fig.caption_zh, font: { name: '宋体' }, size: SZ_5 })] }));
  pars.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: 400, lineRule: 'atLeast', before: 0, after: 200 }, children: [new TextRun({ text: `Figure ${fig.label.replace('图', '')} ${fig.caption_en}`, font: { name: 'Times New Roman' }, size: SZ_5 })] }));
  return pars;
}

// ─── Table insertion ────────────────────────────────────────────────────────

function insertTable(tabKey, registry, inlineData = null) {
  const tab = registry.tables[tabKey];
  if (!tab) { console.warn(`[WARN] Table not found: ${tabKey}`); return []; }
  const rows = inlineData || tab.data || [], cols = tab.columns || [];
  if (rows.length === 0) return [new Paragraph({ spacing: { line: 400, lineRule: 'atLeast' }, children: [new TextRun({ text: `[表格数据缺失: ${tabKey}]`, color: 'FF0000' })] })];

  const pars = [];
  pars.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: 400, lineRule: 'atLeast', before: 200, after: 0 }, children: [new TextRun({ text: `${tab.label} `, font: { name: '宋体' }, size: SZ_5 }), new TextRun({ text: tab.caption_zh, font: { name: '宋体' }, size: SZ_5 })] }));
  pars.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: 400, lineRule: 'atLeast', before: 0, after: 0 }, children: [new TextRun({ text: `Table ${tab.label.replace('表', '')} ${tab.caption_en}`, font: { name: 'Times New Roman' }, size: SZ_5 })] }));

  const ncols = cols.length || (rows[0] ? rows[0].length : 1);
  const cw = Math.floor(CONTENT_WIDTH / ncols), cws = Array(ncols).fill(cw);
  const cb = { style: BorderStyle.SINGLE, size: 6, color: '000000' };
  const trs = [];
  if (cols.length > 0) {
    trs.push(new TableRow({ tableHeader: true, children: cols.map((c, i) => new TableCell({ borders: { top: cb, bottom: cb, left: cb, right: cb }, width: { size: cws[i], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: 400, lineRule: 'atLeast' }, children: mixedTextRuns(String(c), SZ_5, { bold: true }) })] })) }));
  }
  for (const row of rows) {
    trs.push(new TableRow({ children: row.map((c, i) => new TableCell({ borders: { top: cb, bottom: cb, left: cb, right: cb }, width: { size: cws[i], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: 400, lineRule: 'atLeast' }, children: mixedTextRuns(String(c ?? ''), SZ_5) })] })) }));
  }
  pars.push(new Table({ width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: cws, alignment: AlignmentType.CENTER, rows: trs }));
  return pars;
}

// ─── Inline text parser (citations → hyperlinks, bold, inline math) ─────────

function parseInlineText(text, refMap, size = SZ_X4) {
  const tokenRe = /(\[\^ref_[a-zA-Z0-9_]+\])|(\$[^$\n]+?\$)|(\*\*(.+?)\*\*)/g;
  const runs = []; let lastIdx = 0, match;

  while ((match = tokenRe.exec(text)) !== null) {
    if (match.index > lastIdx) runs.push(...mixedTextRuns(text.slice(lastIdx, match.index), size));

    if (match[1]) {
      const rm = match[1].match(/\[\^(ref_[a-zA-Z0-9_]+)\]/);
      if (rm) {
        const refId = rm[1], ref = refMap[refId], ck = ref ? ref.cite_key : '?';
        runs.push(new InternalHyperlink({
          children: [new TextRun({ text: `[${ck}]`, superScript: true, font: { name: 'Times New Roman' }, size })],
          anchor: refId,
        }));
      }
    } else if (match[2]) {
      // Inline math — hidden marker only (replaced by OMML in post-processing)
      const ms = match[2].slice(1, -1);
      const inlKey = `INL_${Object.keys(formulaMap).filter(k => k.startsWith('INL_')).length + 1}`;
      formulaMap[inlKey] = { latex: ms, eqLabel: null };
      // Marker text: hidden (white, tiny). OMML post-processor will find and replace.
      runs.push(new TextRun({ text: `[[${inlKey}]]`, font: { name: 'Times New Roman' }, size: 2, color: 'FFFFFF' }));
      // Fallback: render LaTeX source as tiny grey text (barely visible until OMML replaces it)
      runs.push(new TextRun({ text: ms, font: { name: 'Times New Roman' }, size: SZ_X4, italics: true, color: '888888' }));
    } else if (match[3]) {
      // Bold with inline math processed recursively
              const boldText = match[4];
              const subRe = /\$[^$\n]+?\$/g;
              let subLast = 0, subM;
              while ((subM = subRe.exec(boldText)) !== null) {
                if (subM.index > subLast)
                  runs.push(...mixedTextRuns(boldText.slice(subLast, subM.index), size, { bold: true }));
                const ms = subM[0].slice(1, -1);
                const inlKey = 'INL_' + (Object.keys(formulaMap).filter(k => k.startsWith('INL_')).length + 1);
                formulaMap[inlKey] = { latex: ms, eqLabel: null };
                runs.push(new TextRun({ text: '[[' + inlKey + ']]', font: { name: 'Times New Roman' }, size: 2, color: 'FFFFFF' }));
                runs.push(new TextRun({ text: ms, font: { name: 'Times New Roman' }, size, italics: true, bold: true, color: '888888' }));
                subLast = subM.index + subM[0].length;
              }
              if (subLast < boldText.length)
                runs.push(...mixedTextRuns(boldText.slice(subLast), size, { bold: true }))
    }
    lastIdx = match.index + match[0].length;
  }
  if (lastIdx < text.length) runs.push(...mixedTextRuns(text.slice(lastIdx), size));
  return runs;
}

// ─── Markdown → docx elements ───────────────────────────────────────────────

function parseMarkdown(mdContent, registry, refMap) {
  const elements = [], lines = mdContent.split('\n');
  let i = 0;
  while (i < lines.length) {
    const line = lines[i];

    if (/^<!--\s*(CHAPTER|TITLE|REFS_USED|FIGURES_USED|TABLES_USED):/.test(line)) { i++; continue; }

    // FIGURE marker
    const figM = line.match(/^<!--\s*FIGURE:\s*(\S+)\s*-->/);
    if (figM) { elements.push(...insertFigure(figM[1], registry)); i++; continue; }

    // TABLE marker (inline data)
    const tabO = line.match(/^<!--\s*TABLE:\s*(\S+)$/);
    if (tabO) {
      let idata = null, j = i + 1, dls = [];
      while (j < lines.length && !lines[j].includes('-->')) { dls.push(lines[j]); j++; }
      if (dls.length > 0 && dls[0].trim().startsWith('|'))
        idata = dls.filter(l => l.trim().startsWith('|') && !l.match(/^\|[-\s|]+\|$/)).map(l => l.split('|').slice(1, -1).map(c => c.trim()));
      elements.push(...insertTable(tabO[1], registry, idata));
      i = j + (j < lines.length && lines[j].includes('-->') ? 1 : 0);
      continue;
    }
    const tabS = line.match(/^<!--\s*TABLE:\s*(\S+)\s*-->/);
    if (tabS) { elements.push(...insertTable(tabS[1], registry)); i++; continue; }

    // EQ marker: capture LaTeX + label
    const eqM = line.match(/^<!--\s*EQ:\s*(\S+)\s*-->/);
    if (eqM) {
      const eqLabel = eqM[1];
      // Look ahead/behind for $$ formula
      let latexStr = '';
      if (i + 1 < lines.length && lines[i + 1].trim().startsWith('$$')) {
        latexStr = lines[i + 1].trim().replace(/^\$\$/, '').replace(/\$\$$/, '').trim();
        i += 2;
      } else {
        // Look behind
        let j = i - 1;
        while (j >= 0 && lines[j].trim() === '') j--;
        if (j >= 0 && lines[j].trim().startsWith('$$')) {
          latexStr = lines[j].trim().replace(/^\$\$/, '').replace(/\$\$$/, '').trim();
        }
        i++;
      }
      elements.push(formulaMarker(latexStr, eqLabel));
      continue;
    }

    // Heading
    const h1 = line.match(/^#\s+(.+)/);
    if (h1) { const m = h1[1].match(/^第(\d+)章/); if (m) currentChapter = parseInt(m[1]); elements.push(chapterHeading(h1[1])); i++; continue; }
    const h2 = line.match(/^##\s+(.+)/); if (h2) { elements.push(sectionHeading(h2[1])); i++; continue; }
    const h3 = line.match(/^###\s+(.+)/); if (h3) { elements.push(subsectionHeading(h3[1])); i++; continue; }
    const h4 = line.match(/^####\s+(.+)/); if (h4) { elements.push(subsubsectionHeading(h4[1])); i++; continue; }

    if (line.trim() === '') { i++; continue; }

    // Block formula $$...$$ (single-line or multi-line)
    if (line.trim() === '$$' || (line.trim().startsWith('$$') && line.trim().endsWith('$$'))) {
      let latexLines = [];
      let eqLabel = null;
      let look = 1;

      if (line.trim().startsWith('$$') && line.trim().endsWith('$$') && line.trim().length > 4) {
        // Single-line: $$...$$
        latexLines.push(line.trim().replace(/^\$\$/, '').replace(/\$\$$/, '').trim());
      } else {
        // Multi-line: $$ ... $$
        // Accumulate lines until closing $$
        let j = i + 1;
        while (j < lines.length && !lines[j].trim().endsWith('$$')) {
          latexLines.push(lines[j].trim());
          j++;
        }
        if (j < lines.length && lines[j].trim().endsWith('$$')) {
          latexLines.push(lines[j].trim().replace(/\$\$$/, '').trim());
        }
        look = (j - i) + 1;
      }

      const latexStr = latexLines.join(' ').trim();

      // Look for EQ marker after formula
      while (i + look < lines.length && lines[i + look].trim() === '') look++;
      if (i + look < lines.length) {
        const eqA = lines[i + look].match(/^<!--\s*EQ:\s*(\S+)\s*-->/);
        if (eqA) { eqLabel = eqA[1]; look++; }
      }
      if (!eqLabel) {
        let lb = i - 1;
        while (lb >= 0 && lines[lb].trim() === '') lb--;
        if (lb >= 0) { const eqB = lines[lb].match(/^<!--\s*EQ:\s*(\S+)\s*-->/); if (eqB) eqLabel = eqB[1]; }
      }
      // Only number formulas that have explicit EQ markers
      if (latexStr) {
        elements.push(formulaMarker(latexStr, eqLabel));
      }
      i += look; continue;
    }

    // Quote
    const bq = line.match(/^>\s*(.*)/);
    if (bq) { elements.push(new Paragraph({ alignment: AlignmentType.LEFT, spacing: { line: 400, lineRule: 'atLeast', before: 0, after: 0 }, indent: { left: 480 }, children: parseInlineText(bq[1], refMap, SZ_X4) })); i++; continue; }

    if (/^---+$/.test(line.trim())) { i++; continue; }

    // Markdown list item (starts with - or *)
    const listM = line.match(/^[-*]\s+(.+)/);
    if (listM) {
      elements.push(new Paragraph({
        alignment: AlignmentType.BOTH, spacing: { line: 400, lineRule: 'atLeast', before: 0, after: 0 },
        indent: { left: 480, firstLine: 480 },
        children: parseInlineText(listM[1], refMap, SZ_X4),
      }));
      i++; continue;
    }

    // Body paragraph
    if (line.trim() && !line.startsWith('<!--')) {
      let pt = line.trim(), j = i + 1;
      while (j < lines.length && lines[j].trim() !== '' && !lines[j].startsWith('<!--') && !lines[j].startsWith('#') && !lines[j].startsWith('$$') && !lines[j].startsWith('>') && !/^---+$/.test(lines[j].trim())) { pt += ' ' + lines[j].trim(); j++; }
      elements.push(new Paragraph({ alignment: AlignmentType.BOTH, spacing: { line: 400, lineRule: 'atLeast', before: 0, after: 0 }, indent: { firstLine: 480 }, children: parseInlineText(pt, refMap, SZ_X4) }));
      i = j; continue;
    }
    i++;
  }
  return elements;
}

// ─── Reference section ──────────────────────────────────────────────────────

function buildReferenceSection(refs) {
  const el = [];
  el.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: 400, lineRule: 'atLeast', before: 0, after: 800 }, children: [new TextRun({ text: '参考文献', font: { name: '黑体' }, size: SZ_3 })] }));
  const sorted = [...refs].filter(r => r.cite_key > 0).sort((a, b) => a.cite_key - b.cite_key);
  for (const ref of sorted) {
    el.push(new Paragraph({
      alignment: AlignmentType.BOTH, spacing: { line: 400, lineRule: 'atLeast', before: 0, after: 0 },
      indent: { left: 480, hanging: 480 },
      children: [
        new Bookmark({
          id: ref.id,
          children: [
            new TextRun({ text: `[${ref.cite_key}] `, font: { name: 'Times New Roman' }, size: SZ_X4 }),
            ...mixedTextRuns(ref.formatted || `[${ref.cite_key}] ${ref.title}`, SZ_X4),
          ],
        }),
      ],
    }));
  }
  return el;
}

// ─── Main ───────────────────────────────────────────────────────────────────

async function buildThesis() {
  console.log('Reading registry & references...');
  const registry = JSON.parse(fs.readFileSync(REGISTRY_PATH, 'utf8'));
  const refData = JSON.parse(fs.readFileSync(REFERENCES_PATH, 'utf8'));
  const refMap = {};
  for (const r of refData.refs) refMap[r.id] = r;

  console.log('Reading chapters...');
  const cfiles = fs.readdirSync(CHAPTERS_DIR).filter(f => f.endsWith('.md')).sort();
  const all = [];

  for (let idx = 0; idx < cfiles.length; idx++) {
    const f = cfiles[idx];
    const content = fs.readFileSync(path.join(CHAPTERS_DIR, f), 'utf8');
    console.log(`  Parsing ${f}`);
    if (idx > 0) all.push(new Paragraph({ children: [new PageBreak()] }));
    all.push(...parseMarkdown(content, registry, refMap));
  }

  all.push(new Paragraph({ children: [new PageBreak()] }));
  all.push(...buildReferenceSection(refData.refs));

  console.log('Building DOCX...');
  const doc = new Document({
    styles: { default: { document: { run: { font: { name: '宋体' }, size: SZ_X4 } } } },
    sections: [{
      properties: { page: { size: { width: PAGE_WIDTH, height: PAGE_HEIGHT }, margin: { top: MARGIN_TB, bottom: MARGIN_TB, left: MARGIN_LR, right: MARGIN_LR, header: 794, footer: 1134 } } },
      headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: 400, lineRule: 'atLeast' }, children: [new TextRun({ text: '北京化工大学毕业论文（设计）', font: { name: '宋体' }, size: SZ_X5 })] })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '- ', size: SZ_X5 }), new TextRun({ children: [PageNumber.CURRENT], size: SZ_X5 }), new TextRun({ text: ' -', size: SZ_X5 })] })] }) },
      children: all,
    }],
  });

  const buffer = await Packer.toBuffer(doc);
  fs.mkdirSync(path.dirname(OUTPUT), { recursive: true });
  fs.writeFileSync(OUTPUT, buffer);

  // Write formulas map for Python post-processing
  fs.writeFileSync(FORMULAS_MAP, JSON.stringify(formulaMap, null, 2), 'utf8');
  console.log(`Formulas map: ${Object.keys(formulaMap).length} formulas`);
  console.log(`Done: ${OUTPUT}`);
}

buildThesis().catch(err => { console.error('FAIL:', err); process.exit(1); });
