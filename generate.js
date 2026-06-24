const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, PageNumber, AlignmentType,
  BorderStyle, WidthType, VerticalAlign
} = require('docx');
const fs = require('fs');
const path = require('path');

const procId = process.argv[2];
if (!procId) { console.error('No procedure ID provided'); process.exit(1); }

const dataPath = path.join(__dirname, 'data.json');
if (!fs.existsSync(dataPath)) { console.error('data.json not found'); process.exit(1); }

const data = JSON.parse(fs.readFileSync(dataPath, 'utf8'));
const proc = (data.procs || []).find(p => p.id === procId);
if (!proc) { console.error('Procedure not found:', procId); process.exit(1); }

const users = data.users || [];
const author = users.find(u => u.id === proc.userId);
const opName = proc.op || (author ? author.name : '');
const date = new Date().toLocaleDateString('en-GB');

console.log('Generating document for:', proc.title, '/', proc.ref);

// ─── Parse AI text ────────────────────────────────────────────────────────────
function parseText(txt) {
  const blocks = [];
  (txt || '').replace(/\*\*(.*?)\*\*/g, '$1').replace(/\*(.*?)\*/g, '$1')
    .split('\n').forEach(line => {
      const t = line.trim();
      if (!t) return;
      const clean = t.replace(/^#{1,3}\s*/, '');
      if (/^(I{1,3}|IV|V|VI{1,3}|IX|X)\.\s/i.test(clean))
        blocks.push({ type: 'section', text: clean });
      else if (/^\d+[.)]\s/.test(t))
        blocks.push({ type: 'step', text: t });
      else if (/^[-•]\s/.test(t))
        blocks.push({ type: 'bullet', text: t.slice(2) });
      else if (/^[-=]{3,}$/.test(t)) { /* skip */ }
      else
        blocks.push({ type: 'text', text: clean });
    });
  return blocks;
}

// ─── Image helper ─────────────────────────────────────────────────────────────
function makeImage(ph, w, h) {
  try {
    const mime = ph.data.split(';')[0].split('/')[1]; // jpeg or png
    const b64 = ph.data.split(',')[1];
    const buf = Buffer.from(b64, 'base64');
    return new ImageRun({
      data: buf,
      transformation: { width: w, height: h },
      type: mime === 'jpeg' ? 'jpg' : 'png'
    });
  } catch (e) { return null; }
}

// ─── Cell helpers ─────────────────────────────────────────────────────────────
const B = (style, color) => ({ style: style || BorderStyle.SINGLE, size: 4, color: color || 'AAAAAA' });
const NONE = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders = { top: NONE, bottom: NONE, left: NONE, right: NONE };
const allBorders = { top: B(), bottom: B(), left: B(), right: B() };

function cell(children, width, opts) {
  return new TableCell({
    borders: opts && opts.borders || allBorders,
    width: { size: width, type: WidthType.DXA },
    margins: { top: 80, bottom: 80, left: 140, right: 140 },
    verticalAlign: VerticalAlign.CENTER,
    children
  });
}

// ─── Build document body ──────────────────────────────────────────────────────
const children = [];
const blocks = parseText(proc.genText);
let secCount = -1;
let stepCount = 0;

blocks.forEach(block => {
  if (block.type === 'section') {
    secCount++; stepCount = 0;
    children.push(new Paragraph({
      spacing: { before: 360, after: 120 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '1B3A6B', space: 3 } },
      children: [new TextRun({ text: block.text, bold: true, size: 26, font: 'Arial', color: '1B3A6B' })]
    }));
    return;
  }

  if (block.type === 'step') {
    const sec = proc.sections[secCount];
    const st = sec ? sec.steps[stepCount] : null;
    const photos = (st && st.photos) ? st.photos : [];
    stepCount++;

    if (photos.length) {
      const imgCells = photos.slice(0, 2).map(ph => {
        const img = makeImage(ph, 200, 150);
        return img ? new Paragraph({ alignment: AlignmentType.CENTER, children: [img] }) : null;
      }).filter(Boolean);

      children.push(new Table({
        width: { size: 9638, type: WidthType.DXA },
        columnWidths: [5638, 4000],
        rows: [new TableRow({ children: [
          cell([new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: block.text, size: 22, font: 'Arial' })] })], 5638),
          cell(imgCells.length ? imgCells : [new Paragraph({ children: [] })], 4000)
        ]})]
      }));
      children.push(new Paragraph({ spacing: { before: 60 }, children: [] }));
    } else {
      children.push(new Paragraph({
        indent: { left: 480 },
        spacing: { before: 80, after: 60 },
        children: [new TextRun({ text: block.text, size: 22, font: 'Arial' })]
      }));
    }
    return;
  }

  if (block.type === 'bullet') {
    children.push(new Paragraph({
      indent: { left: 720, hanging: 240 },
      spacing: { before: 40, after: 40 },
      children: [new TextRun({ text: '• ' + block.text, size: 20, font: 'Arial' })]
    }));
    return;
  }

  children.push(new Paragraph({
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text: block.text, size: 20, font: 'Arial', color: '444444' })]
  }));
});

// ─── Photo appendix ───────────────────────────────────────────────────────────
const allPhotos = [];
(proc.sections || []).forEach((sec, si) => {
  (sec.steps || []).forEach((st, idx) => {
    if (st.photos && st.photos.length) allPhotos.push({ sec, si, idx, st });
  });
});

if (allPhotos.length) {
  children.push(new Paragraph({ spacing: { before: 600, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '1B3A6B', space: 3 } },
    children: [new TextRun({ text: 'Photos', bold: true, size: 26, font: 'Arial', color: '1B3A6B' })]
  }));

  allPhotos.forEach(item => {
    children.push(new Paragraph({ spacing: { before: 200, after: 60 },
      children: [new TextRun({ text: item.sec.title + ' — Step ' + (item.idx + 1), bold: true, size: 22, font: 'Arial' })]
    }));
    if (item.st.text) {
      children.push(new Paragraph({ spacing: { before: 0, after: 80 },
        children: [new TextRun({ text: item.st.text.slice(0, 200), size: 20, font: 'Arial', color: '555555' })]
      }));
    }
    const photoParagraphs = [];
    item.st.photos.forEach(ph => {
      const img = makeImage(ph, 320, 240);
      if (img) photoParagraphs.push(new Paragraph({ spacing: { before: 40, after: 40 }, children: [img] }));
    });
    children.push(...photoParagraphs);
  });
}

// ─── Signatures ───────────────────────────────────────────────────────────────
children.push(new Paragraph({ spacing: { before: 480 }, children: [] }));
children.push(new Table({
  width: { size: 9638, type: WidthType.DXA },
  columnWidths: [4819, 4819],
  rows: [
    new TableRow({ children: [
      cell([new Paragraph({ children: [new TextRun({ text: 'Etablit par :', bold: true, size: 20, font: 'Arial' })] })], 4819),
      cell([new Paragraph({ children: [new TextRun({ text: 'Contrôlé par :', bold: true, size: 20, font: 'Arial' })] })], 4819)
    ]}),
    new TableRow({ children: [
      new TableCell({ borders: allBorders, width: { size: 4819, type: WidthType.DXA },
        margins: { top: 400, bottom: 400, left: 140, right: 140 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: opName, size: 22, font: 'Arial' })] })]
      }),
      new TableCell({ borders: allBorders, width: { size: 4819, type: WidthType.DXA },
        margins: { top: 400, bottom: 400, left: 140, right: 140 },
        children: [new Paragraph({ children: [] })]
      })
    ]})
  ]
}));

// ─── Header ──────────────────────────────────────────────────────────────────
const headerTable = new Table({
  width: { size: 9638, type: WidthType.DXA },
  columnWidths: [9638],
  rows: [new TableRow({ children: [
    new TableCell({
      borders: { top: NONE, left: NONE, right: NONE, bottom: { style: BorderStyle.SINGLE, size: 6, color: '1B3A6B', space: 4 } },
      margins: { bottom: 80 },
      children: [
        new Paragraph({ alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: 'Procédure de fabrication', bold: true, size: 28, font: 'Arial' })]
        }),
        new Paragraph({ alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: '\u00ab ' + proc.title + ' \u2013 ' + proc.ref + ' \u00bb', size: 24, font: 'Arial' })]
        })
      ]
    })
  ]})]
});

// ─── Footer ──────────────────────────────────────────────────────────────────
const footerTable = new Table({
  width: { size: 9638, type: WidthType.DXA },
  columnWidths: [5200, 2400, 2038],
  rows: [
    new TableRow({ children: [
      new TableCell({ borders: { top: { style: BorderStyle.SINGLE, size: 4, color: '000000' }, bottom: NONE, left: NONE, right: NONE },
        margins: { top: 60 }, children: [new Paragraph({ children: [new TextRun({ text: 'Doc. : Fabrication ' + proc.ref + ' \u2013 ' + proc.title, size: 16, font: 'Arial' })] })]
      }),
      new TableCell({ borders: { top: { style: BorderStyle.SINGLE, size: 4, color: '000000' }, bottom: NONE, left: NONE, right: NONE },
        margins: { top: 60 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Date : ' + date, size: 16, font: 'Arial' })] })]
      }),
      new TableCell({ borders: { top: { style: BorderStyle.SINGLE, size: 4, color: '000000' }, bottom: NONE, left: NONE, right: NONE },
        margins: { top: 60 }, children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: 'Version :01 / Ed doc :', size: 16, font: 'Arial' })] })]
      })
    ]}),
    new TableRow({ children: [
      new TableCell({ borders: noBorders, children: [new Paragraph({ children: [new TextRun({ text: 'Etablit par : ' + opName, size: 16, font: 'Arial' })] })] }),
      new TableCell({ borders: noBorders, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Contr\u00f4l\u00e9 par :', size: 16, font: 'Arial' })] })] }),
      new TableCell({ borders: noBorders, children: [new Paragraph({ alignment: AlignmentType.RIGHT,
        children: [
          new TextRun({ text: 'Page ', size: 16, font: 'Arial' }),
          new TextRun({ children: [PageNumber.CURRENT], size: 16, font: 'Arial' }),
          new TextRun({ text: ' sur ', size: 16, font: 'Arial' }),
          new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, font: 'Arial' })
        ]
      })] })
    ]})
  ]
});

// ─── Assemble & save ─────────────────────────────────────────────────────────
const doc = new Document({
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1700, right: 1134, bottom: 1800, left: 1134 }
      }
    },
    headers: { default: new Header({ children: [headerTable] }) },
    footers: { default: new Footer({ children: [footerTable] }) },
    children
  }]
});

const outDir = path.join(__dirname, 'docs');
fs.mkdirSync(outDir, { recursive: true });
const outFile = path.join(outDir, procId + '.docx');

Packer.toBuffer(doc)
  .then(buf => {
    fs.writeFileSync(outFile, buf);
    console.log('✓ Generated:', outFile, '(' + Math.round(buf.length/1024) + 'KB)');
  })
  .catch(err => {
    console.error('✗ Error:', err.message);
    process.exit(1);
  });
