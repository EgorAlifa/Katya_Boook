const D = require("./design.js");
const {
  Document, Packer, Paragraph, TextRun, Header, Footer, AlignmentType,
  BorderStyle, PageNumber, PageBreak, Table, TableRow, TableCell, WidthType,
  ShadingType, TabStopType, HeadingLevel, TableOfContents,
  NAVY, INK, GRAY, LIGHT_GRAY, SUBTITLE_GRAY, CALLOUT_TEXT, PAPER,
  FONT_BODY, FONT_MONO,
  PAGE_W, PAGE_H, MARGIN_LR, MARGIN_TOP, MARGIN_BOTTOM, HEADER_H, FOOTER_H,
  CONTENT_W_IN, WORD_STYLES, DEFAULT_STYLES, CHARACTER_STYLES, CHAPTER_NUMBERING, pageProps, makeHeader, makeFooter,
  body, bodyDropCap, subheading, questionHeading, taskLabel, answerLabel, answerText, answerBox,
  calloutBox, chapterOpener, introHeading, figureParagraph, figureRow,
  caption, captionMulti, ruleBreak, formulaParagraph, formulaWhere,
} = D;
const fs = require("fs");
const path = require("path");

// usable body height on a content page (page height minus margins/header/footer), for
// sizing figures the original docx rotated 90/270° to run the full height of the page.
const PAGE_BODY_H_IN = (PAGE_H - MARGIN_TOP - MARGIN_BOTTOM - HEADER_H - FOOTER_H) / 1440;

const blocks = JSON.parse(fs.readFileSync(path.join(__dirname, "classified.json"), "utf8"));

//////////////////////// split into logical sections ////////////////////////
// [0]=intro (before first chapter_mark), [1..3]=chapters, [4]=conclusion(after conclusion_mark)
const segments = [];
let cur = { kind: "intro", num: null, blocks: [] };
for (const b of blocks) {
  if (b.type === "chapter_mark") {
    segments.push(cur);
    const num = (b.text.match(/\d+/) || [""])[0];
    cur = { kind: "chapter", num, blocks: [] };
    continue;
  }
  if (b.type === "conclusion_mark") {
    segments.push(cur);
    cur = { kind: "conclusion", num: null, blocks: [] };
    continue;
  }
  cur.blocks.push(b);
}
segments.push(cur);

//////////////////////// block -> docx content dispatcher ////////////////////////
function renderBlocks(blks) {
  const out = [];
  let titleConsumed = false;

  for (let idx = 0; idx < blks.length; idx++) {
    const b = blks[idx];

    if (b.type === "chapter_title") {
      if (!titleConsumed) {
        titleConsumed = true; // rendered by the caller (chapter opener / intro heading)
        continue;
      }
      out.push(subheading(b.text));
      continue;
    }

    if (b.type === "subheading") {
      // internal subheadings and question-phrased subheadings share the identical REF-Question
      // treatment (navy, left rule) — see design.js subheading()/questionHeading().
      out.push(b.text.trim().endsWith("?") ? questionHeading(b.text) : subheading(b.text));
      continue;
    }

    if (b.type === "task_label") {
      const n = b.text.replace(/ЗАДАЧА\s*/i, "").trim();
      out.push(taskLabel(n));
      continue;
    }

    if (b.type === "answer_label") {
      out.push(answerLabel());
      continue;
    }

    if (b.type === "para_dropcap") {
      out.push(bodyDropCap(b.letter, b.text));
      continue;
    }

    if (b.type === "para") {
      if (b.text && b.text.trim()) out.push(body(b.text));
      continue;
    }

    if (b.type === "formula") {
      out.push(formulaParagraph(b.den1, b.den2));
      out.push(formulaWhere(
        `где: S — синодический период; T${b.den1} и T${b.den2} — периоды обращения соответствующих тел вокруг Солнца`
      ));
      continue;
    }

    if (b.type === "image") {
      if (b.images.length === 1) {
        const im = b.images[0];
        const rotation = im.rotation || 0;
        const sideways = rotation % 180 !== 0;
        const maxH = sideways ? PAGE_BODY_H_IN : (im.h > im.w ? 3.6 : 3.0);
        const p = figureParagraph(im.file, im.w, im.h, undefined, maxH, 160, 60, rotation);
        if (p) out.push(p);
      } else {
        const p = figureRow(b.images, 2.6);
        if (p) out.push(p);
      }
      continue;
    }

    if (b.type === "caption") {
      if (b.items.length === 1) out.push(caption(b.items[0].num, b.items[0].desc));
      else out.push(captionMulti(b.items));
      continue;
    }

    if (b.type === "callout") {
      const bodyLines = b.body_text ? [b.body_text] : [];
      out.push(...calloutBox(b.title || null, bodyLines));
      (b.images || []).forEach((im) => {
        const p = figureParagraph(im.file, im.w, im.h, undefined, 3.4, 160, 60, im.rotation || 0);
        if (p) out.push(p);
      });
      continue;
    }

    if (b.type === "dashtable") {
      out.push(...renderPlanetTable());
      continue;
    }
  }
  return out;
}

//////////////////////// special-cased "Таблица 1" (planet angular drift) ////////////////////////
function renderPlanetTable() {
  const headers = ["", "Меркурий", "Венера", "Земля", "Марс", "Юпитер", "Сатурн", "Уран", "Нептун"];
  const rows = [
    ["Сутки", "4,1", "1,6", "1,0", "0,5", "0,08", "0,03", "0,01", "0,006"],
    ["Месяц", "122,8", "48,0", "29,6", "15,7", "2,5", "1,0", "0,04", "0,18"],
    ["6 месяцев", "736,6", "288,4", "177,4", "94,3", "15,0", "30,0", "2,1", "1,1"],
  ];
  const mkCell = (text, isHeader) =>
    new TableCell({
      shading: isHeader ? { type: ShadingType.CLEAR, color: "auto", fill: NAVY } : undefined,
      margins: { top: 90, bottom: 90, left: 90, right: 90 },
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({
          text, font: isHeader ? FONT_BODY : FONT_MONO, size: 16,
          bold: isHeader, color: isHeader ? "FFFFFF" : INK,
        })],
      })],
    });
  const headerRow = new TableRow({ children: headers.map((h) => mkCell(h, true)) });
  const dataRows = rows.map((r) => new TableRow({ children: r.map((c, i) => mkCell(c, i === 0)) }));
  const tbl = new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [headerRow, ...dataRows] });
  return [
    new Paragraph({ style: "REF-TableLabel", children: [new TextRun({ text: "ТАБЛИЦА 1" })] }),
    tbl,
    caption("", "Угловое смещение планет на фоне звёзд за сутки, месяц и полгода (в градусах)", 220),
  ];
}

//////////////////////// title & copyright pages ////////////////////////
const titleSection = {
  properties: { ...pageProps(), titlePage: true },
  children: [
    new Paragraph({ spacing: { before: 2000 }, children: [new TextRun({ text: "" })] }),
    new Paragraph({ spacing: { before: 200, after: 0 }, alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "НАУЧНО-ПОПУЛЯРНОЕ ИЗДАНИЕ", font: FONT_MONO, size: 16, color: LIGHT_GRAY, characterSpacing: 36 })] }),
    new Paragraph({ style: "REF-Title", spacing: { before: 460, after: 0 }, children: [new TextRun({ text: "Вместе с космонавтами" })] }),
    new Paragraph({ style: "REF-Subtitle", spacing: { before: 200, after: 0 }, children: [new TextRun({ text: "Путешествие по Луне и Марсу" })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 460, after: 0 },
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: NAVY, space: 14 } }, children: [new TextRun({ text: "" })] }),
    new Paragraph({ style: "REF-Authors", spacing: { before: 700, after: 0 },
      children: [new TextRun({ text: "И. А. Ачарова  ·  М. Ю. Невский" })] }),
    new Paragraph({ style: "REF-Technical", spacing: { before: 2600, after: 40 },
      children: [new TextRun({ text: "ООО «Медиа-Полис»" })] }),
    new Paragraph({ style: "REF-Technical", spacing: { before: 0 },
      children: [new TextRun({ text: "Ростов-на-Дону  ·  2026" })] }),
  ],
};

const copyrightSection = {
  properties: { ...pageProps(), type: "nextPage" },
  children: [
    new Paragraph({ spacing: { before: 900 }, children: [new TextRun({ text: "" })] }),
    new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "УДК 52-1  ·  ББК 22.6я721  ·  Н40", font: FONT_MONO, size: 16, color: LIGHT_GRAY })] }),
    new Paragraph({ spacing: { before: 260, after: 120 }, children: [new TextRun({ text: "Ачарова И. А., Невский М. Ю.", bold: true, font: FONT_BODY, size: 24, color: NAVY })] }),
    body("Вместе с космонавтами. Путешествие по Луне и Марсу. — Ростов-н/Д: ООО «Медиа-Полис», 2026.", { spacing: { after: 200, line: 280 } }),
    new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "ISBN: присваивается издательством перед сдачей в печать", italics: true, font: FONT_BODY, size: 20, color: GRAY })] }),
    new Paragraph({ spacing: { after: 260 }, children: [new TextRun({ text: "ББК 22.6я721", font: FONT_MONO, size: 16, color: LIGHT_GRAY })] }),
    body("«Stellarium»: GNU General Public License.", { spacing: { after: 120, line: 280 } }),
    body("Все товарные знаки, бренды третьих лиц, названия продуктов, фирменные наименования и компании, упомянутые в издании, могут быть товарными знаками их соответствующих владельцев и используются в целях обучения и в интересах читателей, не подразумевая нарушения авторского права.", { spacing: { after: 260, line: 280 } }),
    new Paragraph({ spacing: { before: 300 }, border: { top: { style: BorderStyle.SINGLE, size: 4, color: NAVY, space: 10 } },
      children: [new TextRun({ text: "© ООО «Медиа-Полис», 2026", font: FONT_MONO, size: 16, color: LIGHT_GRAY })] }),
  ],
};

//////////////////////// build content sections from parsed segments ////////////////////////
const docSections = [titleSection, copyrightSection];

segments.forEach((seg) => {
  let opener = [];
  let sectionTitle = "";
  if (seg.kind === "intro") {
    sectionTitle = "Введение";
    opener = [introHeading("ВВЕДЕНИЕ")];
  } else if (seg.kind === "chapter") {
    const titleBlock = seg.blocks.find((b) => b.type === "chapter_title");
    const title = titleBlock ? titleBlock.text : ("Глава " + seg.num);
    sectionTitle = "Глава " + seg.num + ". " + title.charAt(0) + title.slice(1).toLowerCase();
    opener = chapterOpener(title);
  } else if (seg.kind === "conclusion") {
    sectionTitle = "Заключение";
    opener = [introHeading("ЗАКЛЮЧЕНИЕ")];
  }

  const body_ = renderBlocks(seg.blocks);
  docSections.push({
    properties: { ...pageProps(), type: "nextPage" },
    headers: { default: makeHeader(sectionTitle) },
    footers: { default: makeFooter() },
    children: [...opener, ...body_],
  });
});

//////////////////////// Table of contents — a REAL Word TOC field ////////////////////////
// Built from Heading 1/2 (docx's TableOfContents -> a native `{ TOC \o "1-2" \h \z \u }` field).
// Word populates/repaginates it itself (features.updateFields below prompts an update on open,
// and it also refreshes normally via References > Update Table or F9) — no more static page-number
// baking or a separate PDF pre-render pass.
const tocSection = {
  properties: { ...pageProps(), type: "nextPage" },
  headers: { default: makeHeader("Содержание") },
  footers: { default: makeFooter() },
  children: [
    introHeading("СОДЕРЖАНИЕ"),
    new TableOfContents("Содержание", { hyperlink: true, headingStyleRange: "1-2" }),
  ],
};

//////////////////////// Colophon (imprint) ////////////////////////
// Real colophon text extracted straight from the source docx's own tail
// (parse_docx.py -> colophon_lines.json) — the old build-full.js had this
// hand-typed with invented placeholder wording; the manuscript's own real
// colophon was sitting unread in parsed_raw.json's ignored "tail" field the
// whole time. Everything here is quoted as printed EXCEPT the typeface line:
// the source says "Гарнитура Times New Roman" (true of the ORIGINAL
// manuscript), which would be a factual error about THIS edition — replaced
// with what this edition actually uses.
const colophonLines = JSON.parse(fs.readFileSync(path.join(__dirname, "colophon_lines.json"), "utf8"));
const grayLine = (text, opts = {}) => new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { after: 20, ...opts },
  children: [new TextRun({ text, font: FONT_MONO, size: 18, color: LIGHT_GRAY })],
});

const colophonSection = {
  properties: { ...pageProps(), type: "nextPage" },
  headers: { default: makeHeader("Колофон") },
  footers: { default: makeFooter() },
  children: [
    new Paragraph({ spacing: { before: 2200 }, children: [new TextRun({ text: "" })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 },
      children: [new TextRun({ text: colophonLines[0], font: FONT_BODY, size: 20, color: INK })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 },
      children: [new TextRun({ text: colophonLines[1], font: FONT_BODY, size: 20, color: INK })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 20 },
      children: [new TextRun({ text: colophonLines[2], bold: true, font: FONT_BODY, size: 24, color: NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 320 },
      children: [new TextRun({ text: colophonLines[3], italics: true, font: FONT_BODY, size: 21, color: GRAY })] }),
    grayLine(colophonLines[4]),   // "Подписано в печать с оригинал-макета ..."
    grayLine(colophonLines[5]),   // "Формат 60×84 1/8. Бумага офсет."
    grayLine("Гарнитуры Literata, JetBrains Mono."), // this edition's real typefaces (source said "Times New Roman" — that was true of the original manuscript, not of this one)
    grayLine(colophonLines[7]),   // "Печать офсетная. Усл. печ. л. ..."
    grayLine(colophonLines[8]),   // "Тираж ... Заказ № ..."
    grayLine(colophonLines[9], { before: 260 }),  // "Отпечатано в типографии ..."
    grayLine(colophonLines[10]),  // address
    grayLine(colophonLines[11]),  // website
  ],
};

//////////////////////// build the document ////////////////////////
const doc = new Document({
  features: { updateFields: true }, // Word updates the TOC + any other fields automatically on open
  styles: { default: DEFAULT_STYLES, paragraphStyles: WORD_STYLES, characterStyles: CHARACTER_STYLES },
  numbering: { config: [CHAPTER_NUMBERING] },
  sections: [...docSections, tocSection, colophonSection],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(path.join(__dirname, "book_full.docx"), buf);
  console.log("written", buf.length, "sections:", docSections.length);
});
