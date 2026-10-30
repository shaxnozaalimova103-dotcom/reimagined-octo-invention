// Ma'ruza matnlarini (manba/*.txt) Word hujjatiga yig'adi.
// Ishga tushirish: node build_docx.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, PageBreak,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle, LevelFormat,
  TableOfContents, Footer, PageNumber,
} = require("docx");

const FONT = "Times New Roman";
const SIZE = 28; // 14 pt
const SRC = path.join(__dirname, "manba");
const OUT = path.join(__dirname, "Kognitiv_tilshunoslik_maruzalar.docx");

// o' g' -> o‘ g‘, qolgan apostroflar -> ’
const typo = (s) => s.replace(/([oOgG])'/g, "$1‘").replace(/'/g, "’");

// **qalin** belgilarini TextRun larga ajratish
function runs(text, extra = {}) {
  return typo(text).split(/(\*\*[^*]+\*\*)/).filter(Boolean).map((part) =>
    part.startsWith("**")
      ? new TextRun({ text: part.slice(2, -2), bold: true, ...extra })
      : new TextRun({ text: part, ...extra }));
}

const body = (text) => new Paragraph({
  children: runs(text), alignment: AlignmentType.JUSTIFIED,
  indent: { firstLine: 709 }, spacing: { after: 80 },
});

const lines = ["01.txt", "02.txt", "03.txt"]
  .flatMap((f) => fs.readFileSync(path.join(SRC, f), "utf8").split("\n"))
  .map((l) => l.trimEnd());

const children = [];
let lecture = 0;
let planOpen = false;

for (const line of lines) {
  if (!line) continue;
  const tag = line.slice(0, 2);
  const rest = line.slice(2);

  if (!line.startsWith("@ ") && planOpen) planOpen = false;

  if (tag === "# ") {
    lecture++;
    if (children.length) children.push(new Paragraph({ children: [new PageBreak()] }));
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: runs(rest) }));
  } else if (line.startsWith("=")) {
    children.push(new Paragraph({
      alignment: AlignmentType.RIGHT, spacing: { after: 120 },
      children: [new TextRun({ text: `Ajratilgan vaqt: ${line.slice(1)}`, italics: true })],
    }));
  } else if (tag === "@ ") {
    if (!planOpen) {
      planOpen = true;
      children.push(new Paragraph({ spacing: { before: 120, after: 60 }, children: [new TextRun({ text: "Reja:", bold: true })] }));
    }
    children.push(new Paragraph({ numbering: { reference: "num", level: 0, instance: lecture * 10 + 1 }, children: runs(rest) }));
  } else if (tag === "##") {
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: runs(line.slice(3)) }));
  } else if (tag === "- ") {
    children.push(new Paragraph({ numbering: { reference: "bullet", level: 0 }, alignment: AlignmentType.JUSTIFIED, spacing: { after: 40 }, children: runs(rest) }));
  } else if (tag === "> ") {
    children.push(new Paragraph({
      alignment: AlignmentType.JUSTIFIED, indent: { left: 709, right: 709 }, spacing: { before: 80, after: 120 },
      border: { left: { style: BorderStyle.SINGLE, size: 12, color: "1F4E79", space: 8 } },
      children: runs(rest, { italics: true }),
    }));
  } else if (tag === "! ") {
    children.push(new Paragraph({
      alignment: AlignmentType.JUSTIFIED, spacing: { before: 200, after: 120 },
      shading: { type: ShadingType.CLEAR, fill: "EAF1F8", color: "auto" },
      children: [new TextRun({ text: "Tayanch tushunchalar: ", bold: true }), ...runs(rest, { italics: true })],
    }));
    children.push(new Paragraph({ spacing: { before: 120, after: 60 }, children: [new TextRun({ text: "Nazorat savollari:", bold: true })] }));
  } else if (tag === "? ") {
    children.push(new Paragraph({ numbering: { reference: "num", level: 0, instance: lecture * 10 + 2 }, children: runs(rest) }));
  } else if (tag === "L ") {
    children.push(new Paragraph({ numbering: { reference: "lit", level: 0 }, alignment: AlignmentType.JUSTIFIED, spacing: { after: 40 }, children: runs(rest) }));
  } else {
    children.push(body(line));
  }
}

// Titul sahifa
const title = [
  new Paragraph({ spacing: { before: 2400 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "«KOGNITIV TILSHUNOSLIK»", bold: true, size: 40 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200 }, children: [new TextRun({ text: "fanidan", size: 32 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 400 }, children: [new TextRun({ text: typo("MA'RUZA MASHG'ULOTLARI UCHUN"), bold: true, size: 32 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: typo("O'QUV MATERIALLARI"), bold: true, size: 32 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 400 }, children: [new TextRun({ text: typo("(15 ta ma'ruza – 30 soat)"), italics: true })] }),
  new Paragraph({ children: [new PageBreak()] }),
];

// Fan mazmuni jadvali
const topics = [
  "Tilshunoslikda antroposentrizmning shakllanishi.",
  "Kognitiv tilshunoslikka kirish (1-qism): kognitiv fan va kognitiv tilshunoslikning paydo bo'lishi.",
  "Kognitiv tilshunoslikka kirish (2-qism): tilning inson va jamiyat hayotidagi o'rni.",
  "Kognitiv tilshunoslikning predmeti va vazifalari. Tafakkurning lisoniy va nolisoniy turlari.",
  "Konsept tafakkur birligi sifatida va uning asosida tushuncha, obraz, lisoniy ma'no umumlashmasi yotishi.",
  "Ssenariy tushunchasi.",
  "Skript tushunchasi.",
  "Freym tushunchasi.",
  "Geshtalt tushunchasi.",
  "Kategoriya tushunchasi. Kategoriyalashning tarixiy va an'anaviy ildizlari.",
  "Lisoniy-kognitiv faoliyatning milliy-madaniy xususiyatlari.",
  "Lison va madaniyat munosabati.",
  "Lisoniy hodisalarning kognitiv tahlili.",
  "Kognitiv tilshunoslikning tayanch atama va tushunchalari izohi (1-qism).",
  "Kognitiv tilshunoslikning tayanch atama va tushunchalari izohi (2-qism).",
];
const W = [800, 7226, 1000];
const border = { style: BorderStyle.SINGLE, size: 4, color: "808080" };
const borders = { top: border, bottom: border, left: border, right: border };
const cell = (text, i, opts = {}) => new TableCell({
  width: { size: W[i], type: WidthType.DXA }, borders,
  margins: { top: 60, bottom: 60, left: 100, right: 100 },
  shading: opts.head ? { type: ShadingType.CLEAR, fill: "D9E2F3", color: "auto" } : undefined,
  children: [new Paragraph({
    alignment: i === 1 ? AlignmentType.LEFT : AlignmentType.CENTER,
    children: [new TextRun({ text: typo(text), bold: !!opts.head, size: 24 })],
  })],
});
const table = new Table({
  width: { size: 9026, type: WidthType.DXA }, columnWidths: W,
  rows: [
    new TableRow({ tableHeader: true, children: [cell("№", 0, { head: true }), cell("Ma'ruza mavzusi", 1, { head: true }), cell("Soat", 2, { head: true })] }),
    ...topics.map((t, i) => new TableRow({ children: [cell(String(i + 1), 0), cell(t, 1), cell("2", 2)] })),
    new TableRow({ children: [cell("", 0, { head: true }), cell("Jami", 1, { head: true }), cell("30", 2, { head: true })] }),
  ],
});

const front = [
  ...title,
  new Paragraph({ heading: HeadingLevel.HEADING_1, children: runs("FAN MAZMUNI (MA'RUZA MASHG'ULOTLARI)") }),
  table,
  new Paragraph({ children: [new PageBreak()] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: "MUNDARIJA", bold: true })] }),
  new TableOfContents("Mundarija", { hyperlink: true, headingStyleRange: "1-1" }),
  new Paragraph({ children: [new PageBreak()] }),
];

const doc = new Document({
  creator: "Kognitiv tilshunoslik",
  title: "Kognitiv tilshunoslik: ma'ruza matnlari",
  features: { updateFields: true },
  styles: {
    default: { document: { run: { font: FONT, size: SIZE }, paragraph: { spacing: { line: 300 } } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 30, bold: true, color: "1F3864" },
        paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 120, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 28, bold: true, color: "1F4E79" },
        paragraph: { spacing: { before: 240, after: 100 }, keepNext: true, outlineLevel: 1 } },
    ],
  },
  numbering: {
    config: [
      { reference: "bullet", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 1069, hanging: 360 } } } }] },
      { reference: "num", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 1069, hanging: 360 } } } }] },
      { reference: "lit", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 567, hanging: 567 } } } }] },
    ],
  },
  sections: [{
    properties: { page: { margin: { top: 1134, bottom: 1134, left: 1701, right: 850 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: [PageNumber.CURRENT], size: 22 })] })] }) },
    children: [...front, ...children],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(OUT, buf);
  console.log("Yozildi:", OUT, `(${lecture} bo'lim)`);
});
