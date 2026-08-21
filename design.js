const fs = require("fs");
const path_ = require("path");
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Header, Footer,
  AlignmentType, BorderStyle, ShadingType, HeightRule, WidthType,
  PageNumber, LevelFormat, convertInchesToTwip, VerticalAlign,
  TabStopType, TabStopPosition, PageBreak, Table, TableRow, TableCell,
  HeadingLevel, TableOfContents, Math: MathZone, MathRun, MathFraction,
} = require("docx");

//////////////////////// DESIGN TOKENS ////////////////////////
// Fixed palette — do not introduce other hexes anywhere in the document.
const NAVY   = "0B1D3A"; // Deep Space — headings, chapter openers
const BLUE   = "1F3A6D"; // secondary accent — rules, "Это интересно", "Ответ"
const ORANGE = "E4572E"; // Mars Orange — sparingly: tasks, questions, small markers only
const PAPER  = "F7F6F2"; // Warm Paper — very light fills only (never the page canvas: books print on
                          // white/cream stock, a solid colour page background wastes ink and looks odd on screen)
const INK    = "17202A"; // Dark Text — body copy (body text is dark, NOT navy)
const GRAY   = "6B7280"; // Muted Gray — hairlines, captions, running heads

const FONT_BODY = "Merriweather";      // body copy only
const FONT_HEAD = "Manrope";           // headings, labels, navigation, captions
const FONT_MONO = "JetBrains Mono";    // formulas, figure/task numbers, page numbers only — never body prose

// Page geometry: the book's OWN colophon states its real print format explicitly —
// "Формат 60×84 1/8" (a standard Russian sheet-fold notation), whose standard trimmed
// page size is ~200x290mm. This is the authoritative source of truth for trim size
// (found in parse_docx.py's colophon extraction — see README) — not a guess derived
// from the Word page setup (which just reflects whatever printer was last selected,
// not the publisher's intended trim) or from an unrelated reference file's size.
const PAGE_W = 11339;
const PAGE_H = 16441;
const MARGIN_LR = 1200;
const MARGIN_TOP = 1200;
const MARGIN_BOTTOM = 1300;
const HEADER_H = 500;
const FOOTER_H = 620;
const CONTENT_W_IN = (PAGE_W - MARGIN_LR * 2) / 1440;

const IMG_DIR = path_.join(__dirname, "img");

//////////////////////// WORD STYLES — real, reusable paragraph styles ////////////////////////
// Every block in the book is written through one of these `style:` ids, never through one-off
// inline formatting — so editing "Heading 1" (or any other style) in Word's Styles pane restyles
// every instance at once, exactly like a normal Word document.
const WORD_STYLES = [
  {
    id: "Title", name: "Title", basedOn: "Normal", next: "Subtitle", quickFormat: true,
    run: { font: FONT_HEAD, bold: true, size: 56, color: NAVY },
    paragraph: { alignment: AlignmentType.CENTER, spacing: { after: 160 } },
  },
  {
    id: "Subtitle", name: "Subtitle", basedOn: "Normal", next: "BodyText", quickFormat: true,
    run: { font: FONT_BODY, italics: true, size: 26, color: BLUE },
    paragraph: { alignment: AlignmentType.CENTER, spacing: { after: 200 } },
  },
  {
    id: "Heading1", name: "heading 1", basedOn: "Normal", next: "FirstParagraph", quickFormat: true,
    run: { font: FONT_HEAD, bold: true, size: 52, color: NAVY },
    paragraph: { spacing: { before: 0, after: 240 }, keepNext: true,
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: GRAY, space: 14 } } },
  },
  {
    id: "Heading2", name: "heading 2", basedOn: "Normal", next: "FirstParagraph", quickFormat: true,
    run: { font: FONT_HEAD, bold: true, size: 28, color: NAVY },
    paragraph: { spacing: { before: 320, after: 160 }, indent: { left: 200 }, keepNext: true, keepLines: true,
      border: { left: { style: BorderStyle.SINGLE, size: 12, color: BLUE, space: 8 } } },
  },
  {
    id: "Heading3", name: "heading 3", basedOn: "Normal", next: "BodyText", quickFormat: true,
    run: { font: FONT_HEAD, bold: true, size: 22, color: BLUE, characterSpacing: 14 },
    paragraph: { spacing: { before: 200, after: 100 }, keepNext: true },
  },
  {
    id: "BodyText", name: "Body Text", basedOn: "Normal", next: "BodyText", quickFormat: true,
    run: { font: FONT_BODY, size: 21, color: INK },
    paragraph: { alignment: AlignmentType.JUSTIFIED, spacing: { after: 140, line: 276 } },
  },
  {
    id: "FirstParagraph", name: "First Paragraph", basedOn: "BodyText", next: "BodyText", quickFormat: true,
    run: { font: FONT_BODY, size: 21, color: INK },
    paragraph: { alignment: AlignmentType.JUSTIFIED, spacing: { after: 140, line: 276 } },
  },
  {
    id: "Caption", name: "Caption", basedOn: "Normal", next: "BodyText", quickFormat: true,
    run: { font: FONT_HEAD, italics: true, size: 18, color: GRAY },
    paragraph: { alignment: AlignmentType.CENTER, spacing: { after: 200 }, keepLines: true },
  },
  {
    id: "Question", name: "Question", basedOn: "Heading2", next: "FirstParagraph", quickFormat: true,
    run: { font: FONT_HEAD, bold: true, size: 26, color: ORANGE },
    paragraph: { spacing: { before: 320, after: 160 }, indent: { left: 200 }, keepNext: true, keepLines: true,
      outlineLevel: 1, // same outline level as Heading 2, so it still lands in the automatic TOC
      border: { left: { style: BorderStyle.SINGLE, size: 12, color: ORANGE, space: 8 } } },
  },
  {
    id: "Task", name: "Task", basedOn: "Normal", next: "BodyText", quickFormat: true,
    run: { font: FONT_HEAD, bold: true, size: 22, color: ORANGE, characterSpacing: 12 },
    paragraph: { spacing: { before: 300, after: 100 }, keepNext: true },
  },
  {
    id: "Answer", name: "Answer", basedOn: "Normal", next: "BodyText", quickFormat: true,
    run: { font: FONT_HEAD, bold: true, size: 22, color: BLUE, characterSpacing: 12 },
    paragraph: { spacing: { before: 220, after: 120 }, keepNext: true,
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: BLUE, space: 10 } } },
  },
  {
    id: "Interesting", name: "Interesting", basedOn: "Normal", next: "Interesting", quickFormat: true,
    run: { font: FONT_BODY, size: 20, color: INK },
    paragraph: { alignment: AlignmentType.JUSTIFIED, spacing: { after: 90, line: 268 } },
  },
  {
    id: "Formula", name: "Formula", basedOn: "Normal", next: "BodyText", quickFormat: true,
    run: { font: FONT_MONO, size: 24, color: NAVY },
    paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 220, after: 220 },
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: BLUE, space: 12 },
                bottom: { style: BorderStyle.SINGLE, size: 4, color: BLUE, space: 12 } } },
  },
  {
    id: "Header", name: "Header", basedOn: "Normal", next: "Header", quickFormat: true,
    run: { font: FONT_HEAD, italics: true, size: 17, color: GRAY },
    paragraph: { spacing: { after: 60 }, border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: GRAY, space: 6 } } },
  },
  {
    id: "Footer", name: "Footer", basedOn: "Normal", next: "Footer", quickFormat: true,
    run: { font: FONT_HEAD, size: 17, color: GRAY },
    paragraph: { spacing: { before: 0 }, border: { top: { style: BorderStyle.SINGLE, size: 2, color: GRAY, space: 6 } } },
  },
];

// Real auto-number for chapters only ("ГЛАВА %1") — attached per-paragraph (not to the Heading1
// style itself, since "Введение"/"Заключение"/"Содержание" are also Heading1 but must NOT be
// numbered as chapters). Word renumbers this field automatically if a chapter is added/removed/reordered.
const CHAPTER_NUMBERING = {
  reference: "chapter-numbering",
  levels: [{
    level: 0, format: LevelFormat.DECIMAL, text: "ГЛАВА %1", start: 1,
    alignment: AlignmentType.LEFT,
    style: { run: { font: FONT_HEAD, bold: true, size: 22, color: BLUE, characterSpacing: 32 } },
  }],
};

//////////////////////// IMAGE HELPERS (native size from the original docx — never
// resized/re-rotated on our own initiative; only downscaled if it would overflow the
// page, and rotated ONLY when the original document itself rotated that picture) ////////////////////////
function imgDimsNative(declWIn, declHIn, maxWidthIn, maxHeightIn, rotation = 0) {
  let wIn = declWIn;
  let hIn = declHIn;
  const swapped = ((rotation % 360) + 360) % 360 % 180 !== 0;
  const footW = swapped ? hIn : wIn;
  const footH = swapped ? wIn : hIn;
  const scale = Math.min(1, maxWidthIn ? maxWidthIn / footW : 1, maxHeightIn ? maxHeightIn / footH : 1);
  if (scale < 1) {
    wIn *= scale;
    hIn *= scale;
  }
  return { width: Math.round(wIn * 96), height: Math.round(hIn * 96) };
}

function imageRunFromFile(file, declWIn, declHIn, maxWidthIn, maxHeightIn, rotation = 0) {
  const path = `${IMG_DIR}/${file}`;
  if (!fs.existsSync(path)) return null;
  const data = fs.readFileSync(path); // raw bytes, unmodified
  const dims = imgDimsNative(declWIn, declHIn, maxWidthIn, maxHeightIn, rotation);
  let ext = (file.split(".").pop() || "jpg").toLowerCase();
  if (ext === "jpeg") ext = "jpg";
  if (!["png", "jpg", "gif", "bmp"].includes(ext)) ext = "jpg";
  return new ImageRun({ type: ext, data, transformation: { ...dims, rotation: rotation || undefined } });
}

// A real inserted picture (ImageRun) in its own paragraph — a plain, robust "Top and Bottom"-style
// placement. (A floating Square/Tight wrap was deliberately avoided: anchored/wrapped images are
// the most common source of "everything shifts when I replace a photo" breakage in Word; a simple
// block-level picture stays put when text before/after it is edited, which matters more here than
// wrap-around text.)
function figureParagraph(file, declWIn, declHIn, maxWidthIn = CONTENT_W_IN, maxHeightIn = 3.4, spacingBefore = 160, spacingAfter = 60, rotation = 0) {
  const run = imageRunFromFile(file, declWIn, declHIn, maxWidthIn, maxHeightIn, rotation);
  if (!run) return null;
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: spacingBefore, after: spacingAfter },
    keepNext: true,
    children: [run],
  });
}

function figureRow(items, maxHeightIn = 2.6, spacingBefore = 160, spacingAfter = 60) {
  const n = items.length;
  const maxWidthIn = CONTENT_W_IN / n - 0.1;
  const runs = [];
  items.forEach((it, idx) => {
    const run = imageRunFromFile(it.file, it.w, it.h, maxWidthIn, maxHeightIn, it.rotation || 0);
    if (run) {
      if (idx > 0) runs.push(new TextRun({ text: "   " }));
      runs.push(run);
    }
  });
  if (!runs.length) return null;
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: spacingBefore, after: spacingAfter },
    keepNext: true,
    children: runs,
  });
}

function ruleBreak(spacingBefore = 220, spacingAfter = 0) {
  return new Paragraph({
    spacing: { before: spacingBefore, after: spacingAfter },
    border: { top: { style: BorderStyle.SINGLE, size: 2, color: GRAY, space: 1 } },
    keepNext: true,
    children: [new TextRun({ text: "" })],
  });
}

// figure/table numbers are the source's own numbering scheme ("Рисунок 2-11", chapter-figure), not
// a simple sequential count — so they stay literal text (see README) rather than a Word SEQ field,
// but the numeral itself is set in JetBrains Mono as "technical/numeric data".
function caption(numLabel, desc, spacingAfter = 200) {
  return new Paragraph({
    style: "Caption",
    spacing: { after: spacingAfter },
    children: [
      new TextRun({ text: numLabel + " ", font: FONT_MONO, italics: false, bold: true, color: BLUE, size: 17 }),
      new TextRun({ text: desc, font: FONT_HEAD, italics: true, color: GRAY, size: 18 }),
    ],
  });
}

function captionMulti(items, spacingAfter = 200) {
  const runs = [];
  items.forEach((it, idx) => {
    if (idx > 0) runs.push(new TextRun({ text: "  ", font: FONT_HEAD, size: 18 }));
    runs.push(new TextRun({ text: it.num + " ", font: FONT_MONO, bold: true, color: BLUE, size: 17 }));
    runs.push(new TextRun({ text: it.desc, font: FONT_HEAD, italics: true, color: GRAY, size: 18 }));
  });
  return new Paragraph({ style: "Caption", spacing: { after: spacingAfter }, children: runs });
}

//////////////////////// TEXT HELPERS ////////////////////////
function body(text, opts = {}) {
  return new Paragraph({
    style: "BodyText",
    children: [new TextRun({ text })],
    ...opts,
  });
}

function bodyDropCap(firstLetter, rest, opts = {}) {
  return new Paragraph({
    style: "FirstParagraph",
    children: [
      new TextRun({ text: firstLetter, bold: true, font: FONT_HEAD, size: 40, color: NAVY }),
      new TextRun({ text: rest }),
    ],
    ...opts,
  });
}

// regular subsection heading (Word Heading 2 — real outline level, appears in the automatic TOC).
// NOTE: `heading:` already applies the built-in "Heading2" paragraph style on its own — passing
// `style:` alongside it would emit a second, duplicate <w:pStyle> and confuse Word/LO's TOC builder
// into skipping the paragraph, so `heading:` is used alone everywhere in this file.
function subheading(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text })] });
}

// a subsection heading phrased as a navigational question — same outline level as subheading (so it
// still shows up in the automatic TOC/navigation pane) but with the "Question" visual treatment.
// Can't use `heading:` here (it would force the built-in "Heading2" style, clobbering "Question"),
// so the outline level is set directly instead — same TOC effect, no duplicate pStyle.
function questionHeading(text) {
  return new Paragraph({
    style: "Question", outlineLevel: 1,
    children: [new TextRun({ text })],
  });
}

function taskLabel(numText) {
  return new Paragraph({
    style: "Task",
    children: [
      new TextRun({ text: "ЗАДАЧА ", font: FONT_HEAD }),
      new TextRun({ text: numText, font: FONT_MONO }),
    ],
  });
}

function box(paragraphs, { fill, borderColor, borderSize = 3, sides = ["top", "bottom", "left", "right"] }) {
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
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [new TableRow({ cantSplit: true, children: [cell] })],
  });
}

function answerLabel() {
  return new Paragraph({ style: "Answer", children: [new TextRun({ text: "ОТВЕТ" })] });
}

function answerText(text, isLast = false) {
  return new Paragraph({
    style: "BodyText",
    spacing: { after: isLast ? 200 : 120 },
    children: [new TextRun({ text })],
  });
}

function answerBox(paragraphs) {
  return [ruleBreak(260, 0), ...paragraphs];
}

function calloutBox(title, bodyLines) {
  const paras = [];
  paras.push(new Paragraph({
    spacing: { after: 90 }, keepNext: true,
    children: [new TextRun({ text: "ЭТО ИНТЕРЕСНО", bold: true, font: FONT_HEAD, size: 17, color: BLUE, characterSpacing: 20 })],
  }));
  if (title) {
    paras.push(new Paragraph({
      spacing: { after: 90 }, keepNext: true,
      children: [new TextRun({ text: title, bold: true, font: FONT_HEAD, size: 21, color: NAVY })],
    }));
  }
  bodyLines.forEach((t, i) => {
    paras.push(new Paragraph({
      style: "Interesting",
      spacing: { after: i === bodyLines.length - 1 ? 0 : 90, line: 268 },
      children: [new TextRun({ text: t })],
    }));
  });
  return [
    box(paras, { fill: "FFFFFF", borderColor: BLUE, borderSize: 3 }),
    new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "" })] }),
  ];
}

// chapter opener: real auto-numbered "ГЛАВА %1" (Word field, per-paragraph numbering — see
// CHAPTER_NUMBERING) + Heading 1 title. Intro/Заключение/Содержание use introHeading() instead
// (same Heading 1 level for the TOC, but never carry the chapter-numbering field).
function chapterOpener(title) {
  return [
    new Paragraph({
      spacing: { before: 0, after: 120 }, keepNext: true,
      numbering: { reference: CHAPTER_NUMBERING.reference, level: 0 },
    }),
    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: title })] }),
  ];
}

function introHeading(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text })],
  });
}

//////////////////////// FORMULA (real, editable Word Equation — not a picture) ////////////////////////
// Renders "1/S = 1/T_<den1> − 1/T_<den2>" as native OMML via docx's Math/MathFraction/MathRun —
// opens in Word's own Equation Editor, fully editable, not a screenshot.
function formulaParagraph(den1, den2) {
  const frac = (num, den) => new MathFraction({ numerator: [new MathRun(num)], denominator: [new MathRun(den)] });
  return new Paragraph({
    style: "Formula",
    children: [
      new MathZone({
        children: [
          frac("1", "S"), new MathRun(" = "),
          frac("1", "T" + den1), new MathRun(" − "),
          frac("1", "T" + den2),
        ],
      }),
    ],
  });
}

function formulaWhere(text) {
  return new Paragraph({
    spacing: { before: 80, after: 200 },
    children: [new TextRun({ text, font: FONT_HEAD, size: 17, color: GRAY, italics: true })],
  });
}

//////////////////////// HEADER / FOOTER (real Word Header/Footer, styled) ////////////////////////
function makeHeader(bookTitle) {
  return new Header({
    children: [
      new Paragraph({
        style: "Header",
        tabStops: [{ type: TabStopType.RIGHT, position: PAGE_W - MARGIN_LR * 2 }],
        children: [
          new TextRun({ text: bookTitle }),
          new TextRun({ text: "\t" }),
          new TextRun({ children: [PageNumber.CURRENT], font: FONT_MONO }),
        ],
      }),
    ],
  });
}

function makeFooter(sectionTitle) {
  return new Footer({
    children: [
      new Paragraph({
        style: "Footer",
        tabStops: [{ type: TabStopType.RIGHT, position: PAGE_W - MARGIN_LR * 2 }],
        children: [
          new TextRun({ text: sectionTitle }),
          new TextRun({ text: "\t" }),
          new TextRun({ children: [PageNumber.CURRENT], bold: true, font: FONT_MONO, size: 18, color: NAVY }),
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
  HeadingLevel, TableOfContents,
  NAVY, BLUE, ORANGE, PAPER, INK, GRAY,
  FONT_BODY, FONT_HEAD, FONT_MONO, PAGE_W, PAGE_H, MARGIN_LR, MARGIN_TOP, MARGIN_BOTTOM,
  HEADER_H, FOOTER_H, CONTENT_W_IN, IMG_DIR,
  WORD_STYLES, CHAPTER_NUMBERING,
  imgDimsNative, imageRunFromFile, figureParagraph, figureRow, ruleBreak,
  caption, captionMulti, body, bodyDropCap, subheading, questionHeading, taskLabel, box,
  answerLabel, answerText, answerBox, calloutBox, chapterOpener, introHeading,
  formulaParagraph, formulaWhere,
  makeHeader, makeFooter, pageProps,
};
