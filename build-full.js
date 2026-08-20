const D = require("./design.js");
const {
  Document, Packer, Paragraph, TextRun, Header, Footer, AlignmentType,
  BorderStyle, PageNumber, PageBreak, Table, TableRow, TableCell, WidthType,
  ShadingType, TabStopType, HeadingLevel, TableOfContents,
  NAVY, BLUE, ORANGE, INK, GRAY, FONT_BODY, FONT_HEAD, FONT_MONO,
  PAGE_W, PAGE_H, MARGIN_LR, MARGIN_TOP, MARGIN_BOTTOM, HEADER_H, FOOTER_H,
  CONTENT_W_IN, WORD_STYLES, CHAPTER_NUMBERING, pageProps, makeHeader, makeFooter,
  body, bodyDropCap, subheading, questionHeading, taskLabel, answerLabel, answerText, answerBox,
  calloutBox, chapterOpener, introHeading, figureParagraph, figureRow,
  caption, captionMulti, ruleBreak, formulaParagraph, formulaWhere,
} = D;
const fs = require("fs");
const path = require("path");

// usable body height on a content page (page height minus margins/header/footer), for
// sizing figures the original docx rotated 90/270° to run the full height of the page.
const PAGE_BODY_H_IN = (PAGE_H - MARGIN_TOP - MARGIN_BOTTOM - HEADER_H - FOOTER_H) / 1440;

const BOOK_TITLE = "Вместе с космонавтами. Путешествие по Луне и Марсу";
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
      // a subheading phrased as a question gets the dedicated "Question" style
      // (still a real Heading 2 — stays in the automatic TOC/navigation pane)
      out.push(b.text.trim().endsWith("?") ? questionHeading(b.text) : subheading(b.text));
      continue;
    }

    if (b.type === "task_label") {
      const n = b.text.replace(/ЗАДАЧА\s*/i, "").trim();
      out.push(taskLabel(n));
      continue;
    }

    if (b.type === "answer_label") {
      out.push(ruleBreak(260, 0));
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
        if (p) out.push(ruleBreak(200, 0)), out.push(p);
      } else {
        const p = figureRow(b.images, 2.6);
        if (p) out.push(ruleBreak(200, 0)), out.push(p);
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
          text, font: isHeader ? FONT_HEAD : FONT_MONO, size: 16,
          bold: isHeader, color: isHeader ? "FFFFFF" : INK,
        })],
      })],
    });
  const headerRow = new TableRow({ children: headers.map((h) => mkCell(h, true)) });
  const dataRows = rows.map((r) => new TableRow({ children: r.map((c, i) => mkCell(c, i === 0)) }));
  const tbl = new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [headerRow, ...dataRows] });
  return [
    ruleBreak(200, 100),
    new Paragraph({ style: "Heading3", children: [new TextRun({ text: "ТАБЛИЦА 1" })] }),
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
      children: [new TextRun({ text: "НАУЧНО-ПОПУЛЯРНОЕ ИЗДАНИЕ", font: FONT_HEAD, size: 16, color: GRAY, characterSpacing: 36 })] }),
    new Paragraph({ style: "Title", spacing: { before: 460, after: 0 }, children: [new TextRun({ text: "Вместе с космонавтами" })] }),
    new Paragraph({ style: "Subtitle", spacing: { before: 200, after: 0 }, children: [new TextRun({ text: "Путешествие по Луне и Марсу" })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 460, after: 0 },
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: GRAY, space: 14 } }, children: [new TextRun({ text: "" })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 700, after: 0 },
      children: [new TextRun({ text: "И. А. Ачарова  ·  М. Ю. Невский", font: FONT_HEAD, size: 22, color: NAVY, characterSpacing: 4 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 2600 },
      children: [new TextRun({ text: "ООО «Медиа-Полис»", font: FONT_HEAD, size: 17, color: GRAY, characterSpacing: 6 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 40 },
      children: [new TextRun({ text: "Ростов-на-Дону  ·  2026", font: FONT_HEAD, size: 17, color: GRAY, characterSpacing: 6 })] }),
  ],
};

const copyrightSection = {
  properties: { ...pageProps(), type: "nextPage" },
  children: [
    new Paragraph({ spacing: { before: 900 }, children: [new TextRun({ text: "" })] }),
    new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "УДК 52-1  ·  ББК 22.6я721  ·  Н40", font: FONT_HEAD, size: 16, color: GRAY })] }),
    new Paragraph({ spacing: { before: 260, after: 120 }, children: [new TextRun({ text: "Ачарова И. А., Невский М. Ю.", bold: true, font: FONT_HEAD, size: 20, color: NAVY })] }),
    body("Вместе с космонавтами. Путешествие по Луне и Марсу. — Ростов-н/Д: ООО «Медиа-Полис», 2026.", { spacing: { after: 200, line: 280 } }),
    new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "ISBN: присваивается издательством перед сдачей в печать", italics: true, font: FONT_BODY, size: 18, color: GRAY })] }),
    new Paragraph({ spacing: { after: 260 }, children: [new TextRun({ text: "ББК 22.6я721", font: FONT_HEAD, size: 16, color: GRAY })] }),
    body("«Stellarium»: GNU General Public License.", { spacing: { after: 120, line: 280 } }),
    body("Все товарные знаки, бренды третьих лиц, названия продуктов, фирменные наименования и компании, упомянутые в издании, могут быть товарными знаками их соответствующих владельцев и используются в целях обучения и в интересах читателей, не подразумевая нарушения авторского права.", { spacing: { after: 260, line: 280 } }),
    new Paragraph({ spacing: { before: 300 }, border: { top: { style: BorderStyle.SINGLE, size: 4, color: GRAY, space: 10 } },
      children: [new TextRun({ text: "© ООО «Медиа-Полис», 2026", font: FONT_HEAD, size: 16, color: GRAY })] }),
  ],
};

//////////////////////// build content sections from parsed segments ////////////////////////
const docSections = [titleSection, copyrightSection];

segments.forEach((seg) => {
  let opener = [];
  let footerTitle = "";
  if (seg.kind === "intro") {
    footerTitle = "Введение";
    opener = [introHeading("ВВЕДЕНИЕ")];
  } else if (seg.kind === "chapter") {
    const titleBlock = seg.blocks.find((b) => b.type === "chapter_title");
    const title = titleBlock ? titleBlock.text : ("Глава " + seg.num);
    footerTitle = "Глава " + seg.num + ". " + title.charAt(0) + title.slice(1).toLowerCase();
    opener = chapterOpener(title);
  } else if (seg.kind === "conclusion") {
    footerTitle = "Заключение";
    opener = [introHeading("ЗАКЛЮЧЕНИЕ")];
  }

  const body_ = renderBlocks(seg.blocks);
  docSections.push({
    properties: { ...pageProps(), type: "nextPage" },
    headers: { default: makeHeader(BOOK_TITLE) },
    footers: { default: makeFooter(footerTitle) },
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
  headers: { default: makeHeader(BOOK_TITLE) },
  footers: { default: makeFooter("Содержание") },
  children: [
    introHeading("СОДЕРЖАНИЕ"),
    new TableOfContents("Содержание", { hyperlink: true, headingStyleRange: "1-2" }),
  ],
};

//////////////////////// Colophon (imprint) ////////////////////////
const colophonSection = {
  properties: { ...pageProps(), type: "nextPage" },
  headers: { default: new Header({ children: [new Paragraph({ children: [] })] }) },
  footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT_MONO, size: 17, color: GRAY })] })] }) },
  children: [
    new Paragraph({ spacing: { before: 2200 }, children: [new TextRun({ text: "" })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 },
      children: [new TextRun({ text: "Ачарова Ирина Александровна", font: FONT_BODY, size: 20, color: INK })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 },
      children: [new TextRun({ text: "Невский Михаил Юрьевич", font: FONT_BODY, size: 20, color: INK })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 20 },
      children: [new TextRun({ text: "Вместе с космонавтами", bold: true, font: FONT_HEAD, size: 24, color: NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 320 },
      children: [new TextRun({ text: "Путешествие по Луне и Марсу", italics: true, font: FONT_BODY, size: 21, color: BLUE })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 20 },
      children: [new TextRun({ text: "Подписано в печать с оригинал-макета — [заполняется издательством]", font: FONT_BODY, size: 18, color: GRAY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 20 },
      children: [new TextRun({ text: "Формат 165×235 мм. Бумага офсетная.", font: FONT_BODY, size: 18, color: GRAY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 20 },
      children: [new TextRun({ text: "Гарнитуры Merriweather, Manrope, JetBrains Mono.", font: FONT_BODY, size: 18, color: GRAY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 20 },
      children: [new TextRun({ text: "Печать офсетная. Усл. печ. л. — [заполняется издательством]", font: FONT_BODY, size: 18, color: GRAY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 20 },
      children: [new TextRun({ text: "Тираж — [заполняется издательством]. Заказ № — [заполняется издательством]", font: FONT_BODY, size: 18, color: GRAY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 260, after: 20 },
      children: [new TextRun({ text: "Отпечатано в типографии ООО «Медиа-Полис»", font: FONT_BODY, size: 18, color: GRAY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 20 },
      children: [new TextRun({ text: "г. Ростов-на-Дону, пр. М. Нагибина, 14 А", font: FONT_BODY, size: 18, color: GRAY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 20 },
      children: [new TextRun({ text: "www.media-polis.ru", font: FONT_BODY, size: 18, color: GRAY })] }),
  ],
};

//////////////////////// build the document ////////////////////////
const doc = new Document({
  features: { updateFields: true }, // Word updates the TOC + any other fields automatically on open
  styles: { paragraphStyles: WORD_STYLES },
  numbering: { config: [CHAPTER_NUMBERING] },
  sections: [...docSections, tocSection, colophonSection],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(path.join(__dirname, "book_full.docx"), buf);
  console.log("written", buf.length, "sections:", docSections.length);
});
