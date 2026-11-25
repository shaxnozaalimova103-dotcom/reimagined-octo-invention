// Oraliq nazorat materiallarini (manba_oraliq/oraliq.txt) Word fayliga yig'adi.
// Belgilar: "# " sarlavha, "P " matn, "## " bo'lim, "S " mavzu kichik sarlavhasi, "? " nazariy savol,
// "A " amaliy topshiriq, "@VARIANTS" yozma ish variantlari jadvali, "Q "/"o "/"K " testlar,
// "|" jadval qatori, "L " adabiyot raqamlari (manba/03.txt dagi umumiy ro'yxat bo'yicha).
// Ishga tushirish: node build_oraliq.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, LevelFormat, Footer, PageNumber,
} = require("docx");

const FONT = "Times New Roman";
const TEXT_W = 9355;
const OUT = path.join(__dirname, "Oraliq_nazorat_1-10-mavzular.docx");

const typo = (s) => s.replace(/([oOgG])'/g, "$1‘").replace(/'/g, "’");
function runs(text, extra = {}) {
  return typo(text).split(/(\*\*[^*]+\*\*)/).filter(Boolean).map((part) =>
    part.startsWith("**")
      ? new TextRun({ text: part.slice(2, -2), bold: true, ...extra })
      : new TextRun({ text: part, ...extra }));
}

const bib = fs.readFileSync(path.join(__dirname, "manba", "03.txt"), "utf8")
  .split("\n").filter((l) => l.startsWith("L ")).map((l) => l.slice(2).trim());
const extra = {
  G2: "Koffka K. Principles of Gestalt Psychology. – New York: Harcourt, Brace and Company, 1935.",
};
const ref = (k) => {
  const r = /^\d+$/.test(k) ? bib[Number(k) - 1] : extra[k];
  if (!r) throw new Error(`adabiyot topilmadi: ${k}`);
  return r;
};

const lines = fs.readFileSync(path.join(__dirname, "manba_oraliq", "oraliq.txt"), "utf8")
  .split("\n").map((l) => l.trimEnd()).filter(Boolean);
const theory = lines.filter((l) => l.startsWith("? ")).map((l) => l.slice(2));
const practice = lines.filter((l) => l.startsWith("A ")).map((l) => l.slice(2));
if (theory.length !== practice.length * 2) throw new Error(`nazariy ${theory.length}, amaliy ${practice.length}`);

const border = { style: BorderStyle.SINGLE, size: 4, color: "808080" };
function table(rows, widths) {
  widths = widths || (() => {
    const w = Math.floor(TEXT_W / rows[0].length);
    return rows[0].map((_, i) => (i === rows[0].length - 1 ? TEXT_W - w * (rows[0].length - 1) : w));
  })();
  return new Table({
    width: { size: TEXT_W, type: WidthType.DXA }, columnWidths: widths,
    rows: rows.map((r, ri) => new TableRow({
      tableHeader: ri === 0, cantSplit: true,
      children: widths.map((cw, ci) => new TableCell({
        width: { size: cw, type: WidthType.DXA },
        borders: { top: border, bottom: border, left: border, right: border },
        margins: { top: 40, bottom: 40, left: 80, right: 80 },
        shading: ri === 0 ? { type: ShadingType.CLEAR, fill: "D9E2F3", color: "auto" } : undefined,
        children: [new Paragraph({ alignment: ci === 0 ? AlignmentType.CENTER : AlignmentType.LEFT,
          children: runs(r[ci] || "", { size: 22, bold: ri === 0 }) })],
      })),
    })),
  });
}

// Variant i: nazariy i, nazariy (n + i), amaliy i
function variants() {
  const n = practice.length;
  const box = { style: BorderStyle.SINGLE, size: 6, color: "1F4E79", space: 4 };
  const boxed = (children, extra = {}) => new Paragraph({
    border: { top: extra.top ? box : undefined, bottom: extra.bottom ? box : undefined, left: box, right: box },
    shading: extra.top ? { type: ShadingType.CLEAR, fill: "D9E2F3", color: "auto" } : undefined,
    alignment: extra.top ? AlignmentType.CENTER : AlignmentType.JUSTIFIED,
    indent: extra.top ? undefined : { left: 340, hanging: 340 },
    keepNext: !extra.bottom, keepLines: true, spacing: { after: extra.bottom ? 240 : 40 }, children,
  });
  const out = [];
  for (let i = 0; i < n; i++) {
    const qs = [theory[i], theory[n + i], practice[i]];
    out.push(boxed([new TextRun({ text: `${i + 1}-VARIANT`, bold: true })], { top: true }));
    qs.forEach((q, j) => out.push(boxed(runs(`${j + 1}. ${q}`), { bottom: j === 2 })));
  }
  return out;
}

const out = [];
let tableRows = [];
const flush = () => {
  if (!tableRows.length) return;
  out.push(table(tableRows), new Paragraph({ spacing: { after: 120 }, children: [] }));
  tableRows = [];
};
for (const line of lines) {
  if (line.startsWith("|")) { tableRows.push(line.slice(1).split("|").map((c) => c.trim())); continue; }
  flush();
  const tag = line.slice(0, 2), rest = line.slice(2);
  if (tag === "# ") out.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: runs(rest) }));
  else if (tag === "##") out.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: runs(line.slice(3)) }));
  else if (line === "@VARIANTS") out.push(...variants());
  else if (tag === "S ") out.push(new Paragraph({ spacing: { before: 160, after: 60 }, keepNext: true,
    children: runs(rest, { bold: true, italics: true }) }));
  else if (tag === "? ") out.push(new Paragraph({ numbering: { reference: "theory", level: 0 }, alignment: AlignmentType.JUSTIFIED,
    spacing: { after: 40 }, children: runs(rest) }));
  else if (tag === "A ") out.push(new Paragraph({ numbering: { reference: "practice", level: 0 }, alignment: AlignmentType.JUSTIFIED,
    spacing: { after: 40 }, children: runs(rest) }));
  else if (tag === "Q ") out.push(new Paragraph({ numbering: { reference: "test", level: 0 }, spacing: { before: 100, after: 20 },
    keepNext: true, children: runs(rest) }));
  else if (tag === "o ") out.push(new Paragraph({ indent: { left: 1069 }, spacing: { after: 0 }, children: runs(rest) }));
  else if (tag === "K ") out.push(new Paragraph({ spacing: { before: 160 }, children: [new TextRun({ text: "Testlar kaliti: ", bold: true }), ...runs(rest)] }));
  else if (tag === "L ") {
    out.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: runs("VI. Tavsiya etiladigan adabiyotlar") }));
    for (const k of rest.split(",").map((x) => x.trim())) {
      out.push(new Paragraph({ numbering: { reference: "lit", level: 0 }, alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 40 }, children: runs(ref(k)) }));
    }
  } else if (tag === "P ") out.push(new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 80 }, children: runs(rest) }));
  else throw new Error(`noma'lum qator: ${line}`);
}
flush();

const numCfg = (reference, left, hanging) =>
  ({ reference, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left, hanging } } } }] });

const doc = new Document({
  title: typo("Kognitiv tilshunoslik: oraliq nazorat"),
  styles: {
    default: { document: { run: { font: FONT, size: 28 }, paragraph: { spacing: { line: 300 } } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 30, bold: true, color: "1F3864" },
        paragraph: { alignment: AlignmentType.CENTER, spacing: { after: 240 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 28, bold: true, color: "1F4E79" },
        paragraph: { spacing: { before: 300, after: 120 }, keepNext: true, outlineLevel: 1 } },
    ],
  },
  numbering: { config: [numCfg("theory", 709, 567), numCfg("practice", 709, 567), numCfg("test", 709, 567), numCfg("lit", 567, 567)] },
  sections: [{
    properties: { page: { margin: { top: 1134, bottom: 1134, left: 1701, right: 850 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: [PageNumber.CURRENT], size: 22 })] })] }) },
    children: out,
  }],
});
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(OUT, buf); console.log("Yozildi:", OUT); });

// Biletlar: har bir biletda umumiy tushunchaga oid bitta nazariy savol (manba_oraliq/biletlar.txt)
const TICKETS_OUT = path.join(__dirname, "Oraliq_nazorat_biletlar.docx");
const ticketQuestions = fs.readFileSync(path.join(__dirname, "manba_oraliq", "biletlar.txt"), "utf8")
  .split("\n").filter((l) => l.startsWith("? ")).map((l) => l.slice(2).trim());
const tBox = { style: BorderStyle.SINGLE, size: 8, color: "1F4E79", space: 6 };
const tPara = (children, { first, last, align = AlignmentType.CENTER, after = 40 } = {}) => new Paragraph({
  border: { top: first ? tBox : undefined, bottom: last ? tBox : undefined, left: tBox, right: tBox },
  alignment: align, keepNext: !last, keepLines: true, spacing: { after: last ? 360 : after }, children,
});
const tickets = ticketQuestions.flatMap((q, i) => [
  tPara([new TextRun({ text: "________________________________ universiteti", size: 22 })], { first: true }),
  tPara(runs("«Kognitiv tilshunoslik» fanidan oraliq nazorat", { bold: true, size: 24 })),
  tPara([new TextRun({ text: `BILET № ${i + 1}`, bold: true, size: 28, color: "1F3864" })], { after: 120 }),
  tPara(runs(`Savol: ${q}`), { align: AlignmentType.JUSTIFIED, after: 200 }),
  tPara(runs("Kafedra mudiri: ________________          Fan o'qituvchisi: ________________", { size: 22 }), { last: true }),
]);
const ticketDoc = new Document({
  title: typo("Kognitiv tilshunoslik: oraliq nazorat biletlari"),
  styles: { default: { document: { run: { font: FONT, size: 28 }, paragraph: { spacing: { line: 276 } } } } },
  sections: [{
    properties: { page: { margin: { top: 851, bottom: 851, left: 1418, right: 851 } } },
    children: tickets,
  }],
});
Packer.toBuffer(ticketDoc).then((buf) => { fs.writeFileSync(TICKETS_OUT, buf); console.log("Yozildi:", TICKETS_OUT); });
