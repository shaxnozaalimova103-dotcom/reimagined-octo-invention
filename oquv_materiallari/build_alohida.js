// Har bir ma'ruzani alohida Word fayliga yig'adi (maruzalar/ papkasiga).
// Tartib: mavzu → Reja → Tayanch tushunchalar → ma'ruza matni → Nazorat savollari → Foydalanilgan adabiyotlar.
// Ishga tushirish: node build_alohida.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  ShadingType, BorderStyle, LevelFormat, Footer, PageNumber,
} = require("docx");

const FONT = "Times New Roman";
const SRC = path.join(__dirname, "manba");
const OUT_DIR = path.join(__dirname, "maruzalar");

const typo = (s) => s.replace(/([oOgG])'/g, "$1‘").replace(/'/g, "’");
function runs(text, extra = {}) {
  return typo(text).split(/(\*\*[^*]+\*\*)/).filter(Boolean).map((part) =>
    part.startsWith("**")
      ? new TextRun({ text: part.slice(2, -2), bold: true, ...extra })
      : new TextRun({ text: part, ...extra }));
}

// Matnni ma'ruzalarga ajratish
const lines = ["01.txt", "02.txt", "03.txt"]
  .flatMap((f) => fs.readFileSync(path.join(SRC, f), "utf8").split("\n"))
  .map((l) => l.trimEnd()).filter(Boolean);
const blocks = [];
for (const l of lines) {
  if (l.startsWith("# ")) blocks.push({ title: l.slice(2), plan: [], terms: "", body: [], questions: [], lit: [] });
  else {
    const b = blocks[blocks.length - 1];
    if (l.startsWith("@ ")) b.plan.push(l.slice(2));
    else if (l.startsWith("! ")) b.terms = l.slice(2);
    else if (l.startsWith("? ")) b.questions.push(l.slice(2));
    else if (l.startsWith("L ")) b.lit.push(l.slice(2));
    else if (!l.startsWith("=")) b.body.push(l);
  }
}
const bib = blocks.pop().lit; // umumiy ro'yxat (1-dan raqamlangan)
const extra = {
  G1: "Lakoff G. Linguistic Gestalts // Papers from the Thirteenth Regional Meeting of the Chicago Linguistic Society. – Chicago, 1977. – P. 236–287.",
  G2: "Koffka K. Principles of Gestalt Psychology. – New York: Harcourt, Brace and Company, 1935.",
};
const ref = (k) => (typeof k === "number" ? bib[k - 1] : extra[k]);

// Har bir ma'ruza uchun adabiyotlar (umumiy ro'yxatdagi raqamlar) va fayl nomi
const meta = [
  ["Antroposentrizm", [1, 3, 4, 16, 17, 18, 8, 9]],
  ["Kognitiv_tilshunoslikka_kirish_1", [1, 3, 7, 8, 9, 10, 13, 33]],
  ["Kognitiv_tilshunoslikka_kirish_2", [1, 3, 19, 30, 9, 31, 44, 45]],
  ["Predmet_va_vazifalar_Tafakkur_turlari", [1, 7, 8, 9, 10, 19, 13, 14]],
  ["Konsept", [1, 5, 20, 21, 22, 9, 32, 10]],
  ["Ssenariy", [1, 9, 7, 24, 39, 10, 13, 23]],
  ["Skript", [1, 9, 7, 24, 39, 10, 15, 13]],
  ["Freym", [1, 9, 12, 23, 38, 7, 13, 46]],
  ["Geshtalt", [1, 9, 7, "G1", "G2", 36, 13, 33]],
  ["Kategoriya_va_kategoriyalash", [1, 26, 34, 40, 9, 12, 15, 14]],
  ["Milliy-madaniy_xususiyatlar", [1, 3, 6, 21, 27, 42, 30, 10]],
  ["Lison_va_madaniyat", [5, 6, 11, 28, 29, 31, 30, 22]],
  ["Kognitiv_tahlil", [1, 25, 33, 37, 35, 43, 9, 13]],
  ["Tayanch_atamalar_1", [1, 5, 7, 9, 8, 13, 14, 18]],
  ["Tayanch_atamalar_2", [1, 5, 7, 41, 35, 36, 28, 29]],
];
if (meta.length !== blocks.length) throw new Error(`ma'ruzalar soni ${blocks.length}, meta ${meta.length}`);

const label = (text, center = false) => new Paragraph({
  alignment: center ? AlignmentType.CENTER : AlignmentType.LEFT,
  spacing: { before: 240, after: 100 }, keepNext: true,
  children: [new TextRun({ text: typo(text), bold: true })],
});
const listItem = (ref, text, just = false) => new Paragraph({
  numbering: { reference: ref, level: 0 }, spacing: { after: 40 },
  alignment: just ? AlignmentType.JUSTIFIED : AlignmentType.LEFT, children: runs(text),
});

function bodyPara(line) {
  const tag = line.slice(0, 2), rest = line.slice(2);
  if (tag === "##") return new Paragraph({ heading: HeadingLevel.HEADING_2, children: runs(line.slice(3)) });
  if (tag === "- ") return listItem("bullet", rest, true);
  if (tag === "> ") return new Paragraph({
    alignment: AlignmentType.JUSTIFIED, indent: { left: 709, right: 709 }, spacing: { before: 80, after: 120 },
    border: { left: { style: BorderStyle.SINGLE, size: 12, color: "1F4E79", space: 8 } },
    children: runs(rest, { italics: true }),
  });
  return new Paragraph({ children: runs(line), alignment: AlignmentType.JUSTIFIED, indent: { firstLine: 709 }, spacing: { after: 80 } });
}

const numCfg = (reference, left, hanging, format = LevelFormat.DECIMAL, text = "%1.") =>
  ({ reference, levels: [{ level: 0, format, text, alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left, hanging } } } }] });

fs.mkdirSync(OUT_DIR, { recursive: true });
Promise.all(blocks.map((b, i) => {
  const [slug, litKeys] = meta[i];
  const children = [
    new Paragraph({ heading: HeadingLevel.HEADING_1, children: runs(b.title) }),
    label("Reja:", true),
    ...b.plan.map((p) => listItem("plan", p)),
    new Paragraph({
      alignment: AlignmentType.JUSTIFIED, spacing: { before: 240, after: 200 },
      shading: { type: ShadingType.CLEAR, fill: "EAF1F8", color: "auto" },
      children: [new TextRun({ text: "Tayanch tushunchalar: ", bold: true }), ...runs(b.terms, { italics: true })],
    }),
    ...b.body.map(bodyPara),
    label("Nazorat savollari:"),
    ...b.questions.map((q) => listItem("questions", q)),
    label("Foydalanilgan adabiyotlar:"),
    ...litKeys.map((k) => {
      const r = ref(k);
      if (!r) throw new Error(`adabiyot topilmadi: ${k}`);
      return listItem("lit", r, true);
    }),
  ];
  const doc = new Document({
    title: typo(b.title),
    styles: {
      default: { document: { run: { font: FONT, size: 28 }, paragraph: { spacing: { line: 300 } } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: FONT, size: 30, bold: true, color: "1F3864" },
          paragraph: { alignment: AlignmentType.CENTER, spacing: { after: 120 }, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: FONT, size: 28, bold: true, color: "1F4E79" },
          paragraph: { spacing: { before: 240, after: 100 }, keepNext: true, outlineLevel: 1 } },
      ],
    },
    numbering: { config: [
      numCfg("plan", 1069, 360), numCfg("questions", 1069, 360), numCfg("lit", 567, 567),
      numCfg("bullet", 1069, 360, LevelFormat.BULLET, "•"),
    ] },
    sections: [{
      properties: { page: { margin: { top: 1134, bottom: 1134, left: 1701, right: 850 } } },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
        children: [new TextRun({ children: [PageNumber.CURRENT], size: 22 })] })] }) },
      children,
    }],
  });
  const file = path.join(OUT_DIR, `${String(i + 1).padStart(2, "0")}-maruza_${slug}.docx`);
  return Packer.toBuffer(doc).then((buf) => fs.writeFileSync(file, buf));
})).then(() => console.log(`${blocks.length} ta fayl yozildi: ${OUT_DIR}`));
