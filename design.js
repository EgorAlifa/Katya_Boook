const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Header, Footer,
  AlignmentType, BorderStyle, ShadingType, HeightRule, WidthType,
  PageNumber, LevelFormat, convertInchesToTwip, VerticalAlign,
  TabStopType, TabStopPosition, PageBreak, Table, TableRow, TableCell,
} = require("docx");

//////////////////////// DESIGN TOKENS — restrained academic style ////////////////////////
const NAVY      = "1F3864";
const NAVY_SOFT = "44577A";
const RULE      = "C7CCD6";
const RULE_SOFT = "DCE0E7";
const CALLOUT_BG    = "F4F6F9";
const CALLOUT_BORDER= "CBD2DE";
const GRAY_TXT  = "6B7280";
const INK       = "1A1A1A";

const FONT_BODY = "Georgia";
const FONT_LABEL = "Arial";

// Page geometry: book trim 165 x 235 mm
const PAGE_W = 9354;
const PAGE_H = 13323;
const MARGIN_LR = 1000;
const MARGIN_TOP = 1000;
const MARGIN_BOTTOM = 1100;
const HEADER_H = 500;
const FOOTER_H = 620;
const CONTENT_W_IN = (PAGE_W - MARGIN_LR * 2) / 1440;

const IMG_DIR = "/home/claude/work/img";

//////////////////////// IMAGE HELPERS (use pandoc-declared inch sizes) ////////////////////////
function imgDimsFromDeclared(declWIn, declHIn, widthFrac, maxHeightIn) {
  let wIn = CONTENT_W_IN * widthFrac;
  let hIn = wIn * (declHIn / declWIn);
  if (maxHeightIn && hIn > maxHeightIn) {
    hIn = maxHeightIn;
    wIn = hIn * (declWIn / declHIn);
  }
  return { width: Math.round(wIn * 96), height: Math.round(hIn * 96) };
}

function imageRunFromFile(file, declWIn, declHIn, widthFrac, maxHeightIn) {
  const path = `${IMG_DIR}/${file}`;
  if (!fs.existsSync(path)) return null;
  const data = fs.readFileSync(path);
  const dims = imgDimsFromDeclared(declWIn, declHIn, widthFrac, maxHeightIn);
  let ext = (file.split(".").pop() || "jpg").toLowerCase();
  if (ext === "jpg") ext = "jpg";
  if (ext === "jpeg") ext = "jpg";
  if (!["png", "jpg", "gif", "bmp"].includes(ext)) ext = "jpg";
  return new ImageRun({ type: ext, data, transformation: dims });
}

// single figure, centred, capped to a fraction of content width / max height
function figureParagraph(file, declWIn, declHIn, widthFrac = 0.7, maxHeightIn = 3.4, spacingBefore = 160, spacingAfter = 60) {
  const run = imageRunFromFile(file, declWIn, declHIn, widthFrac, maxHeightIn);
  if (!run) return null;
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: spacingBefore, after: spacingAfter },
    children: [run],
  });
}

// several images side-by-side in one row (e.g. two comparison figures)
function figureRow(items, maxHeightIn = 2.6, spacingBefore = 160, spacingAfter = 60) {
  const n = items.length;
  const frac = Math.min(0.92 / n, 0.48);
  const runs = [];
  items.forEach((it, idx) => {
    const run = imageRunFromFile(it.file, it.w, it.h, frac, maxHeightIn);
    if (run) {
      if (idx > 0) runs.push(new TextRun({ text: "   " }));
      runs.push(run);
    }
  });
  if (!runs.length) return null;
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: spacingBefore, after: spacingAfter },
    children: runs,
  });
}

function ruleBreak(spacingBefore = 220, spacingAfter = 0) {
  return new Paragraph({
    spacing: { before: spacingBefore, after: spacingAfter },
    border: { top: { style: BorderStyle.SINGLE, size: 4, color: RULE, space: 1 } },
    children: [new TextRun({ text: "" })],
  });
}

function caption(numLabel, desc, spacingAfter = 200) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: spacingAfter },
    children: [
      new TextRun({ text: numLabel + " ", bold: true, italics: true, color: NAVY_SOFT, font: FONT_BODY, size: 19 }),
      new TextRun({ text: desc, italics: true, color: GRAY_TXT, font: FONT_BODY, size: 19 }),
    ],
  });
}

// multiple caption items combined into one paragraph, e.g. "Рисунок 1-4. ... Рисунок 1-5. ..."
function captionMulti(items, spacingAfter = 200) {
  const runs = [];
  items.forEach((it, idx) => {
    if (idx > 0) runs.push(new TextRun({ text: "  ", font: FONT_BODY, size: 19 }));
    runs.push(new TextRun({ text: it.num + " ", bold: true, italics: true, color: NAVY_SOFT, font: FONT_BODY, size: 19 }));
    runs.push(new TextRun({ text: it.desc, italics: true, color: GRAY_TXT, font: FONT_BODY, size: 19 }));
  });
  return new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: spacingAfter }, children: runs });
}

//////////////////////// TEXT HELPERS ////////////////////////
function body(text, opts = {}) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { after: 140, line: 288 },
    children: [new TextRun({ text, font: FONT_BODY, size: 21, color: INK })],
    ...opts,
  });
}

function bodyDropCap(firstLetter, rest, opts = {}) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { after: 140, line: 288 },
    children: [
      new TextRun({ text: firstLetter, bold: true, font: FONT_BODY, size: 40, color: NAVY }),
      new TextRun({ text: rest, font: FONT_BODY, size: 21, color: INK }),
    ],
    ...opts,
  });
}

function subheading(text) {
  return new Paragraph({
    spacing: { before: 320, after: 160 },
    indent: { left: 200 },
    border: { left: { style: BorderStyle.SINGLE, size: 18, color: NAVY, space: 8 } },
    children: [new TextRun({ text, bold: true, font: FONT_BODY, size: 24, color: NAVY })],
  });
}

function taskLabel(text) {
  return new Paragraph({
    spacing: { before: 300, after: 100 },
    children: [
      new TextRun({ text: text.toUpperCase(), bold: true, font: FONT_LABEL, size: 18, color: NAVY_SOFT, characterSpacing: 18 }),
    ],
  });
}

function box(paragraphs, { fill, borderColor, borderSize = 6, sides = ["top", "bottom", "left", "right"] }) {
  const spec = { style: BorderStyle.SINGLE, size: borderSize, color: borderColor };
  const borders = {};
  sides.forEach((s) => (borders[s] = spec));
  ["top", "bottom", "left", "right"].forEach((s) => {
    if (!borders[s]) borders[s] = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  });
  const cell = new TableCell({
    children: paragraphs,
    shading: { type: ShadingType.CLEAR, color: "auto", fill },
    borders,
    margins: { top: 160, bottom: 160, left: 220, right: 220 },
    width: { size: 100, type: WidthType.PERCENTAGE },
  });
  return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [new TableRow({ children: [cell] })] });
}

function answerLabel() {
  return new Paragraph({
    spacing: { before: 220, after: 100 },
    children: [new TextRun({ text: "ОТВЕТ", bold: true, font: FONT_LABEL, size: 18, color: NAVY_SOFT, characterSpacing: 18 })],
  });
}

function answerText(text, isLast = false) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { after: isLast ? 200 : 120, line: 300 },
    children: [new TextRun({ text, font: FONT_BODY, size: 21, color: INK })],
  });
}

function answerBox(paragraphs) {
  return [ruleBreak(260, 0), ...paragraphs];
}

function calloutBox(title, bodyLines) {
  const paras = [];
  paras.push(new Paragraph({
    spacing: { after: 90 },
    children: [new TextRun({ text: "ЭТО ИНТЕРЕСНО", bold: true, font: FONT_LABEL, size: 17, color: NAVY_SOFT, characterSpacing: 20 })],
  }));
  if (title) {
    paras.push(new Paragraph({
      spacing: { after: 90 },
      children: [new TextRun({ text: title, bold: true, italics: true, font: FONT_BODY, size: 21, color: NAVY })],
    }));
  }
  bodyLines.forEach((t, i) => {
    paras.push(new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { after: i === bodyLines.length - 1 ? 0 : 90, line: 280 },
      children: [new TextRun({ text: t, font: FONT_BODY, size: 20, color: INK })],
    }));
  });
  return [
    box(paras, { fill: CALLOUT_BG, borderColor: CALLOUT_BORDER, borderSize: 4 }),
    new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "" })] }),
  ];
}

function chapterOpener(num, title) {
  const kicker = num ? "ГЛАВА " + num : null;
  const arr = [];
  if (kicker) {
    arr.push(new Paragraph({
      spacing: { before: 0, after: 120 },
      children: [new TextRun({ text: kicker, bold: true, font: FONT_LABEL, size: 19, color: NAVY_SOFT, characterSpacing: 32 })],
    }));
  }
  arr.push(new Paragraph({
    spacing: { before: 0, after: 240 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: RULE, space: 14 } },
    children: [new TextRun({ text: title, bold: true, font: FONT_BODY, size: 38, color: NAVY })],
  }));
  return arr;
}

function introHeading(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 320 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: RULE, space: 12 } },
    children: [new TextRun({ text, bold: true, font: FONT_BODY, size: 32, color: NAVY, characterSpacing: 20 })],
  });
}

//////////////////////// HEADER / FOOTER ////////////////////////
function makeHeader(bookTitle) {
  return new Header({
    children: [
      new Paragraph({
        tabStops: [{ type: TabStopType.RIGHT, position: PAGE_W - MARGIN_LR * 2 }],
        spacing: { after: 60 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: RULE, space: 6 } },
        children: [
          new TextRun({ text: bookTitle, italics: true, font: FONT_BODY, size: 17, color: GRAY_TXT }),
          new TextRun({ text: "\t", font: FONT_BODY }),
          new TextRun({ children: [PageNumber.CURRENT], font: FONT_BODY, size: 17, color: GRAY_TXT }),
        ],
      }),
    ],
  });
}

function makeFooter(sectionTitle) {
  return new Footer({
    children: [
      new Paragraph({
        tabStops: [{ type: TabStopType.RIGHT, position: PAGE_W - MARGIN_LR * 2 }],
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: RULE, space: 6 } },
        children: [
          new TextRun({ text: sectionTitle, font: FONT_BODY, size: 17, color: GRAY_TXT }),
          new TextRun({ text: "\t", font: FONT_BODY }),
          new TextRun({ children: [PageNumber.CURRENT], bold: true, font: FONT_BODY, size: 18, color: NAVY }),
        ],
      }),
    ],
  });
}

const pageProps = () => ({
  page: {
    size: { width: PAGE_W, height: PAGE_H },
    margin: { top: MARGIN_TOP, bottom: MARGIN_BOTTOM, left: MARGIN_LR, right: MARGIN_LR, header: HEADER_H, footer: FOOTER_H },
  },
});

module.exports = {
  fs, Document, Packer, Paragraph, TextRun, ImageRun, Header, Footer,
  AlignmentType, BorderStyle, ShadingType, HeightRule, WidthType,
  PageNumber, LevelFormat, convertInchesToTwip, VerticalAlign,
  TabStopType, TabStopPosition, PageBreak, Table, TableRow, TableCell,
  NAVY, NAVY_SOFT, RULE, RULE_SOFT, CALLOUT_BG, CALLOUT_BORDER, GRAY_TXT, INK,
  FONT_BODY, FONT_LABEL, PAGE_W, PAGE_H, MARGIN_LR, MARGIN_TOP, MARGIN_BOTTOM,
  HEADER_H, FOOTER_H, CONTENT_W_IN, IMG_DIR,
  imgDimsFromDeclared, imageRunFromFile, figureParagraph, figureRow, ruleBreak,
  caption, captionMulti, body, bodyDropCap, subheading, taskLabel, box,
  answerLabel, answerText, answerBox, calloutBox, chapterOpener, introHeading,
  makeHeader, makeFooter, pageProps,
};
