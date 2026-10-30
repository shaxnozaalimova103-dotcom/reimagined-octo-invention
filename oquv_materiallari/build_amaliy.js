// Amaliy mashg'ulotlarni (manba_amaliy/*.txt) alohida Word fayllarga yig'adi (amaliy/ papkasiga).
// Belgilar: "# " mavzu, "M " maqsad, "## " bo'lim, "T " topshiriq sarlavhasi, "- " ro'yxat,
// "N " namuna, "|" jadval qatori, "Q " test savoli, "o " javob varianti, "K " test kaliti,
// "? " nazariy savol, "L " adabiyot raqamlari (manba/03.txt dagi umumiy ro'yxat bo'yicha).
// Ishga tushirish: node build_amaliy.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, LevelFormat, Footer, PageNumber,
} = require("docx");

const FONT = "Times New Roman";
const SRC = path.join(__dirname, "manba_amaliy");
const OUT_DIR = path.join(__dirname, "amaliy");
const TEXT_W = 9355; // A4 kengligi minus chap (3 sm) va o'ng (1,5 sm) hoshiyalar

const typo = (s) => s.replace(/([oOgG])'/g, "$1‘").replace(/'/g, "’");
function runs(text, extra = {}) {
  return typo(text).split(/(\*\*[^*]+\*\*)/).filter(Boolean).map((part) =>
    part.startsWith("**")
      ? new TextRun({ text: part.slice(2, -2), bold: true, ...extra })
      : new TextRun({ text: part, ...extra }));
}

// Umumiy adabiyotlar ro'yxati (ma'ruzalar bilan bir xil raqamlash) va qo'shimcha manbalar
const bib = fs.readFileSync(path.join(__dirname, "manba", "03.txt"), "utf8")
  .split("\n").filter((l) => l.startsWith("L ")).map((l) => l.slice(2).trim());
const extra = {
  G1: "Lakoff G. Linguistic Gestalts // Papers from the Thirteenth Regional Meeting of the Chicago Linguistic Society. – Chicago, 1977. – P. 236–287.",
  G2: "Koffka K. Principles of Gestalt Psychology. – New York: Harcourt, Brace and Company, 1935.",
  K1: "Кун Т. Структура научных революций. – М.: Прогресс, 1975.",
  S1: "Сепир Э. Избранные труды по языкознанию и культурологии. – М.: Прогресс, 1993.",
  W1: "Уорф Б.Л. Отношение норм поведения и мышления к языку // Новое в лингвистике. Вып. 1. – М.: Изд-во иностранной литературы, 1960.",
  J1: "Жинкин Н.И. Речь как проводник информации. – М.: Наука, 1982.",
  P1: "Pinker S. The Language Instinct: How the Mind Creates Language. – New York: William Morrow, 1994.",
  B1: "Бабушкин А.П. Типы концептов в лексико-фразеологической семантике языка. – Воронеж: Изд-во ВГУ, 1996.",
  F1: "Фрумкина Р.М. Концептуальный анализ с точки зрения лингвиста и психолога // Научно-техническая информация. Сер. 2. – 1992. – №3.",
  Z1: "Залевская А.А. Психолингвистический подход к проблеме концепта // Методологические проблемы когнитивной лингвистики. – Воронеж: Изд-во ВГУ, 2001.",
  BT1: "Bartlett F.C. Remembering: A Study in Experimental and Social Psychology. – Cambridge: Cambridge University Press, 1932.",
};
const ref = (k) => {
  const r = /^\d+$/.test(k) ? bib[Number(k) - 1] : extra[k];
  if (!r) throw new Error(`adabiyot topilmadi: ${k}`);
  return r;
};

const slugs = [
  "Antroposentrizm_1", "Antroposentrizm_2", "Kognitiv_tilshunoslik_orni", "Atama_Tafakkur_UPK",
  "Konsept_izohlari_1", "Konsept_izohlari_2", "Ssenariy_va_slotlar", "Skript",
  "Geshtalt_va_yaxlitlik", "Freym_olti_bolagi", "Kategoriya", "Til_va_madaniyat",
  "Kognitiv_tahlil", "Tayanch_atamalar_1", "Tayanch_atamalar_2",
];

// Matnni mashg'ulotlarga ajratish
const sessions = [];
for (const f of ["a1.txt", "a2.txt", "a3.txt"]) {
  for (const l of fs.readFileSync(path.join(SRC, f), "utf8").split("\n").map((x) => x.trimEnd())) {
    if (!l) continue;
    if (l.startsWith("# ")) sessions.push({ title: l.slice(2), lines: [] });
    else sessions[sessions.length - 1].lines.push(l);
  }
}
if (sessions.length !== slugs.length) throw new Error(`mashg'ulotlar soni: ${sessions.length}`);

const border = { style: BorderStyle.SINGLE, size: 4, color: "808080" };
function table(rows) {
  const cols = rows[0].length;
  const w = Math.floor(TEXT_W / cols);
  const widths = Array.from({ length: cols }, (_, i) => (i === cols - 1 ? TEXT_W - w * (cols - 1) : w));
  return new Table({
    width: { size: TEXT_W, type: WidthType.DXA }, columnWidths: widths,
    rows: rows.map((r, ri) => new TableRow({
      tableHeader: ri === 0,
      children: widths.map((cw, ci) => new TableCell({
        width: { size: cw, type: WidthType.DXA },
        borders: { top: border, bottom: border, left: border, right: border },
        margins: { top: 40, bottom: 40, left: 80, right: 80 },
        shading: ri === 0 ? { type: ShadingType.CLEAR, fill: "D9E2F3", color: "auto" } : undefined,
        children: [new Paragraph({ children: runs(r[ci] || "", { size: 22, bold: ri === 0 }) })],
      })),
    })),
  });
}

function build(s, idx) {
  const out = [new Paragraph({ heading: HeadingLevel.HEADING_1, children: runs(s.title) })];
  let tableRows = [];
  const flush = () => {
    if (!tableRows.length) return;
    out.push(table(tableRows), new Paragraph({ spacing: { after: 60 }, children: [] }));
    tableRows = [];
  };
  for (const line of s.lines) {
    if (line.startsWith("|")) { tableRows.push(line.slice(1).split("|").map((c) => c.trim())); continue; }
    flush();
    const tag = line.slice(0, 2), rest = line.slice(2);
    if (tag === "M ") {
      out.push(new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 160 },
        children: [new TextRun({ text: typo("Mashg'ulot maqsadi: "), bold: true }), ...runs(rest, { italics: true })] }));
    } else if (tag === "##") {
      out.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: runs(line.slice(3)) }));
    } else if (tag === "T ") {
      out.push(new Paragraph({ spacing: { before: 200, after: 80 }, keepNext: true, alignment: AlignmentType.JUSTIFIED,
        children: runs(rest, { bold: true }) }));
    } else if (tag === "- ") {
      out.push(new Paragraph({ numbering: { reference: "bullet", level: 0 }, alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 40 }, children: runs(rest) }));
    } else if (tag === "N ") {
      out.push(new Paragraph({ alignment: AlignmentType.JUSTIFIED, indent: { left: 284 }, spacing: { after: 40 },
        shading: { type: ShadingType.CLEAR, fill: "F2F2F2", color: "auto" },
        border: { left: { style: BorderStyle.SINGLE, size: 12, color: "2E75B6", space: 6 } },
        children: runs(rest) }));
    } else if (tag === "Q ") {
      out.push(new Paragraph({ numbering: { reference: "test", level: 0, instance: idx }, spacing: { before: 100, after: 20 },
        keepNext: true, children: runs(rest) }));
    } else if (tag === "o ") {
      out.push(new Paragraph({ indent: { left: 1069 }, spacing: { after: 0 }, children: runs(rest) }));
    } else if (tag === "K ") {
      out.push(new Paragraph({ spacing: { before: 160 }, children: [new TextRun({ text: "Testlar kaliti: ", bold: true }), ...runs(rest)] }));
    } else if (tag === "? ") {
      out.push(new Paragraph({ numbering: { reference: "theory", level: 0, instance: idx }, alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 40 }, children: runs(rest) }));
    } else if (tag === "L ") {
      out.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: runs("Foydalanilgan adabiyotlar") }));
      for (const k of rest.split(",").map((x) => x.trim())) {
        out.push(new Paragraph({ numbering: { reference: "lit", level: 0, instance: idx }, alignment: AlignmentType.JUSTIFIED,
          spacing: { after: 40 }, children: runs(ref(k)) }));
      }
    } else {
      out.push(new Paragraph({ alignment: AlignmentType.JUSTIFIED, indent: { firstLine: 709 }, spacing: { after: 80 }, children: runs(line) }));
    }
  }
  flush();
  return out;
}

const numCfg = (reference, left, hanging, format = LevelFormat.DECIMAL, text = "%1.") =>
  ({ reference, levels: [{ level: 0, format, text, alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left, hanging } } } }] });

fs.mkdirSync(OUT_DIR, { recursive: true });
Promise.all(sessions.map((s, i) => {
  const doc = new Document({
    title: typo(s.title),
    styles: {
      default: { document: { run: { font: FONT, size: 28 }, paragraph: { spacing: { line: 300 } } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: FONT, size: 30, bold: true, color: "1F3864" },
          paragraph: { alignment: AlignmentType.CENTER, spacing: { after: 200 }, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: FONT, size: 28, bold: true, color: "1F4E79" },
          paragraph: { spacing: { before: 280, after: 120 }, keepNext: true, outlineLevel: 1 } },
      ],
    },
    numbering: { config: [
      numCfg("test", 709, 360), numCfg("theory", 1069, 360), numCfg("lit", 567, 567),
      numCfg("bullet", 1069, 360, LevelFormat.BULLET, "•"),
    ] },
    sections: [{
      properties: { page: { margin: { top: 1134, bottom: 1134, left: 1701, right: 850 } } },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
        children: [new TextRun({ children: [PageNumber.CURRENT], size: 22 })] })] }) },
      children: build(s, i + 1),
    }],
  });
  const file = path.join(OUT_DIR, `${String(i + 1).padStart(2, "0")}-amaliy_${slugs[i]}.docx`);
  return Packer.toBuffer(doc).then((buf) => fs.writeFileSync(file, buf));
})).then(() => console.log(`${sessions.length} ta fayl yozildi: ${OUT_DIR}`));
