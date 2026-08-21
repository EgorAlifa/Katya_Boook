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
// Fixed palette per the layout-correction brief — no orange anywhere, a single
// navy accent everywhere headings/labels/rules used to be split between navy/
// blue/orange, exact grays for secondary/technical text.
const NAVY         = "1B2952"; // primary accent — headings, dropcap, ЗАДАЧА/ОТВЕТ, rules, page number
const INK           = "232118"; // body copy
const GRAY          = "666666"; // captions, header running text, secondary
const LIGHT_GRAY    = "999999"; // light monospace/technical text (title-page imprint line)
const SUBTITLE_GRAY = "54534B"; // title-page subtitle only
const CALLOUT_TEXT  = "3A372F"; // "Это интересно" body text
const PAPER         = "F7F6F2"; // very light neutral fill — "Это интересно" block background only,
                                 // never the page canvas (books print on white/cream stock)

const FONT_BODY = "Literata";              // every heading/body/caption/title in the book
const FONT_MONO = "JetBrains Mono Medium"; // ЗАДАЧА/ОТВЕТ, "ЭТО ИНТЕРЕСНО" label, page numbers,
                                            // title-page technical imprint line — never body prose

// Page geometry: A4 portrait, ~2cm margins on every side (matches the brief's "около 2 см"
// and yields the ~17cm text column the reference PDF's layout implies).
const PAGE_W = 11906;      // 210mm
const PAGE_H = 16838;      // 297mm
const MARGIN_LR = 1134;    // ~2cm
const MARGIN_TOP = 1134;
const MARGIN_BOTTOM = 1134;
const HEADER_H = 680;
const FOOTER_H = 680;
const CONTENT_W_IN = (PAGE_W - MARGIN_LR * 2) / 1440;

const IMG_DIR = path_.join(__dirname, "img");

//////////////////////// WORD STYLES — real, reusable paragraph styles ////////////////////////
// Every block in the book is written through one of these `style:` ids, never through one-off
// inline formatting — so editing a style in Word's Styles pane restyles every instance at once.
// All custom styles are named "REF-*" per the layout-correction brief; built-in Heading1/Heading2
// keep their functional ids (required for `heading:` + the automatic TOC to work) but their
// *display names* are renamed to the mandated "REF-Chapter" / "REF-Question".
const WORD_STYLES = [
  {
    id: "REF-Title", name: "REF-Title", basedOn: "Normal", next: "REF-Subtitle", quickFormat: true,
    run: { font: FONT_BODY, bold: true, size: 64, color: NAVY },
    paragraph: { alignment: AlignmentType.CENTER, spacing: { after: 160 } },
  },
  {
    id: "REF-Subtitle", name: "REF-Subtitle", basedOn: "Normal", next: "REF-Body", quickFormat: true,
    run: { font: FONT_BODY, italics: true, size: 34, color: SUBTITLE_GRAY },
    paragraph: { alignment: AlignmentType.CENTER, spacing: { after: 200 } },
  },
  {
    id: "REF-Authors", name: "REF-Authors", basedOn: "Normal", next: "REF-Technical", quickFormat: true,
    run: { font: FONT_BODY, size: 28, color: NAVY },
    paragraph: { alignment: AlignmentType.CENTER, spacing: { after: 0 } },
  },
  {
    id: "REF-Technical", name: "REF-Technical", basedOn: "Normal", next: "REF-Technical", quickFormat: true,
    run: { font: FONT_MONO, size: 22, color: LIGHT_GRAY },
    paragraph: { alignment: AlignmentType.CENTER, spacing: { after: 40 } },
  },
  {
    id: "REF-Body", name: "REF-Body", basedOn: "Normal", next: "REF-Body", quickFormat: true,
    run: { font: FONT_BODY, size: 28, color: INK },
    paragraph: { alignment: AlignmentType.JUSTIFIED, spacing: { after: 140, line: 276 } },
  },
  {
    id: "REF-FirstParagraph", name: "REF-FirstParagraph", basedOn: "REF-Body", next: "REF-Body", quickFormat: true,
    run: { font: FONT_BODY, size: 28, color: INK },
    paragraph: { alignment: AlignmentType.JUSTIFIED, spacing: { after: 140, line: 276 } },
  },
  {
    id: "REF-Caption", name: "REF-Caption", basedOn: "Normal", next: "REF-Body", quickFormat: true,
    run: { font: FONT_BODY, italics: true, size: 24, color: GRAY },
    paragraph: { alignment: AlignmentType.CENTER, spacing: { after: 200 }, keepLines: true },
  },
  {
    id: "REF-Task", name: "REF-Task", basedOn: "Normal", next: "REF-Body", quickFormat: true,
    run: { font: FONT_MONO, bold: false, size: 24, color: NAVY, characterSpacing: 10 },
    paragraph: { spacing: { before: 300, after: 160 }, keepNext: true,
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: NAVY, space: 8 } } },
  },
  {
    id: "REF-Answer", name: "REF-Answer", basedOn: "Normal", next: "REF-Body", quickFormat: true,
    run: { font: FONT_MONO, bold: false, size: 24, color: NAVY, characterSpacing: 10 },
    paragraph: { spacing: { before: 300, after: 160 }, keepNext: true,
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: NAVY, space: 8 } } },
  },
  {
    id: "REF-Info", name: "REF-Info", basedOn: "Normal", next: "REF-Info", quickFormat: true,
    run: { font: FONT_BODY, size: 26, color: CALLOUT_TEXT },
    paragraph: { alignment: AlignmentType.JUSTIFIED, spacing: { after: 90, line: 268 } },
  },
  {
    id: "REF-Formula", name: "REF-Formula", basedOn: "Normal", next: "REF-Body", quickFormat: true,
    run: { font: FONT_MONO, size: 24, color: NAVY },
    paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 220, after: 220 } },
  },
  {
    id: "REF-TableLabel", name: "REF-TableLabel", basedOn: "Normal", next: "REF-Body", quickFormat: true,
    run: { font: FONT_BODY, bold: true, size: 26, color: NAVY, characterSpacing: 16 },
    paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 200, after: 100 }, keepNext: true },
  },
  {
    id: "REF-Header", name: "REF-Header", basedOn: "Normal", next: "REF-Header", quickFormat: true,
    run: { font: FONT_BODY, italics: true, size: 22, color: GRAY },
    paragraph: { spacing: { after: 0 } },
  },
];

// Heading1/Heading2 are overridden here (via `styles.default`), NOT as entries in WORD_STYLES —
// docx.js always emits its own built-in Heading1/Heading2/Title/etc. style definitions from
// DefaultStylesFactory regardless of what's in `paragraphStyles`; adding a same-id entry there
// produces TWO <w:style styleId="Heading1"> elements in styles.xml (confirmed by inspecting the
// generated XML), and LibreOffice resolves that collision by using the FIRST (unstyled default,
// blue #2E74B5, left-aligned) one instead of ours. `styles.default.heading1/heading2` customizes
// the factory's own single definition in place — this is the correct API for overriding built-in
// styles that `heading:` refers to (`name` here is what shows up in Word's Styles pane, so this is
// also where "REF-Chapter"/"REF-Question" are set, same as any other renamed style).
const DEFAULT_STYLES = {
  heading1: {
    name: "REF-Chapter", quickFormat: true,
    run: { font: FONT_BODY, bold: true, size: 40, color: NAVY },
    paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 0, after: 260 }, keepNext: true },
  },
  heading2: {
    name: "REF-Question", quickFormat: true,
    run: { font: FONT_BODY, bold: true, size: 34, color: NAVY },
    paragraph: { spacing: { before: 320, after: 160 }, indent: { left: 200 }, keepNext: true, keepLines: true,
      border: { left: { style: BorderStyle.SINGLE, size: 16, color: NAVY, space: 8 } } },
  },
};

// A genuine Word CHARACTER style (not a paragraph style) — applied only to the page-number
// run inside the header, so it renders as JetBrains Mono Medium/navy against the surrounding
// Literata italic gray section name, per the brief's mandated "REF-PageNumber" style.
const CHARACTER_STYLES = [
  {
    id: "REF-PageNumber", name: "REF-PageNumber", basedOn: "DefaultParagraphFont",
    run: { font: FONT_MONO, size: 22, color: NAVY },
  },
];

// Real auto-number for chapters only ("ГЛАВА %1") — attached per-paragraph (not to the Heading1
// style itself, since "Введение"/"Заключение"/"Содержание" are also Heading1 but must NOT be
// numbered as chapters). Word renumbers this field automatically if a chapter is added/removed/reordered.
const CHAPTER_NUMBERING = {
  reference: "chapter-numbering",
  levels: [{
    level: 0, format: LevelFormat.DECIMAL, text: "ГЛАВА %1", start: 1,
    alignment: AlignmentType.CENTER,
    style: { run: { font: FONT_BODY, bold: true, size: 24, color: NAVY, characterSpacing: 28 } },
  }],
};

//////////////////////// IMAGE HELPERS (native size from the original docx — never
// resized/re-rotated on our own initiative; only downscaled if it would overflow the
// page, and rotated ONLY when the original document itself rotated that picture) ////////////////////////
// The layout brief caps any deviation from the image's own declared size at ±15% — so the
// overflow-fit scale is clamped to that band instead of shrinking arbitrarily to fit the page.
const MIN_SCALE = 0.85;

function imgDimsNative(declWIn, declHIn, maxWidthIn, maxHeightIn, rotation = 0) {
  let wIn = declWIn;
  let hIn = declHIn;
  const swapped = ((rotation % 360) + 360) % 360 % 180 !== 0;
  const footW = swapped ? hIn : wIn;
  const footH = swapped ? wIn : hIn;
  let scale = Math.min(1, maxWidthIn ? maxWidthIn / footW : 1, maxHeightIn ? maxHeightIn / footH : 1);
  if (scale < MIN_SCALE) scale = MIN_SCALE; // never shrink an image more than 15% below its native size
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

// figure/table numbers use the source's own numbering scheme ("Рисунок 2-11", chapter-figure), so
// they stay literal text (see README) rather than a Word SEQ field. Per the layout brief the whole
// caption is one plain Literata Italic gray run directly under the image — no bold/mono/color accent.
function caption(numLabel, desc, spacingAfter = 200) {
  return new Paragraph({
    style: "REF-Caption",
    spacing: { after: spacingAfter },
    children: [new TextRun({ text: (numLabel ? numLabel + " " : "") + desc })],
  });
}

function captionMulti(items, spacingAfter = 200) {
  const text = items.map((it) => (it.num ? it.num + " " : "") + it.desc).join("   ");
  return new Paragraph({ style: "REF-Caption", spacing: { after: spacingAfter }, children: [new TextRun({ text })] });
}

//////////////////////// TEXT HELPERS ////////////////////////
function body(text, opts = {}) {
  return new Paragraph({
    style: "REF-Body",
    children: [new TextRun({ text })],
    ...opts,
  });
}

function bodyDropCap(firstLetter, rest, opts = {}) {
  return new Paragraph({
    style: "REF-FirstParagraph",
    children: [
      new TextRun({ text: firstLetter, bold: true, font: FONT_BODY, size: 64, color: NAVY }),
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

// A subsection heading phrased as a navigational question. The layout brief gives internal
// subheadings and questions the identical visual treatment (navy, left rule) — so this is the
// exact same built-in Heading2 style as subheading(); kept as a separate named function only so
// call sites stay self-documenting about which kind of heading they're rendering.
function questionHeading(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text })] });
}

function taskLabel(numText) {
  return new Paragraph({
    style: "REF-Task",
    children: [new TextRun({ text: "ЗАДАЧА " + numText })],
  });
}

function answerLabel() {
  return new Paragraph({ style: "REF-Answer", children: [new TextRun({ text: "ОТВЕТ" })] });
}

function answerText(text, isLast = false) {
  return new Paragraph({
    style: "REF-Body",
    spacing: { after: isLast ? 200 : 120 },
    children: [new TextRun({ text })],
  });
}

function answerBox(paragraphs) {
  return [...paragraphs];
}

// "Это интересно" callout — a plain shaded block (no table/border/frame): a light neutral fill
// carried across every paragraph in the block with matching indent, so it reads as one seamless
// panel. Padding top/bottom is a thin same-fill spacer paragraph rather than a box margin.
function calloutBox(title, bodyLines) {
  const fill = { type: ShadingType.CLEAR, color: "auto", fill: PAPER };
  const pad = { left: 220, right: 220 };
  const paras = [];
  paras.push(new Paragraph({ shading: fill, indent: pad, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: "" })] }));
  paras.push(new Paragraph({
    shading: fill, indent: pad, spacing: { after: 90 }, keepNext: true,
    children: [new TextRun({ text: "ЭТО ИНТЕРЕСНО", font: FONT_MONO, size: 22, color: NAVY, characterSpacing: 20 })],
  }));
  if (title) {
    paras.push(new Paragraph({
      shading: fill, indent: pad, spacing: { after: 90 }, keepNext: true,
      children: [new TextRun({ text: title, italics: true, font: FONT_BODY, size: 30, color: NAVY })],
    }));
  }
  bodyLines.forEach((t, i) => {
    paras.push(new Paragraph({
      style: "REF-Info", shading: fill, indent: pad,
      spacing: { after: i === bodyLines.length - 1 ? 0 : 90, line: 268 },
      children: [new TextRun({ text: t })],
    }));
  });
  paras.push(new Paragraph({ shading: fill, indent: pad, spacing: { before: 0, after: 0 }, children: [new TextRun({ text: "" })] }));
  paras.push(new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "" })] }));
  return paras;
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
    children: [new TextRun({ text })],
  });
}

//////////////////////// FORMULA (real, editable Word Equation — not a picture) ////////////////////////
// Renders "1/S = 1/T_<den1> − 1/T_<den2>" as native OMML via docx's Math/MathFraction/MathRun —
// opens in Word's own Equation Editor, fully editable, not a screenshot.
function formulaParagraph(den1, den2) {
  const frac = (num, den) => new MathFraction({ numerator: [new MathRun(num)], denominator: [new MathRun(den)] });
  return new Paragraph({
    style: "REF-Formula",
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
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 200 },
    children: [new TextRun({ text, font: FONT_BODY, size: 22, color: GRAY, italics: true })],
  });
}

//////////////////////// HEADER / FOOTER (real Word Header/Footer) ////////////////////////
// Header (top of page only): current section name on the left, automatic page number on the
// right — nothing else, no book title. Footer: completely empty, by explicit request.
function makeHeader(sectionTitle) {
  return new Header({
    children: [
      new Paragraph({
        style: "REF-Header",
        tabStops: [{ type: TabStopType.RIGHT, position: PAGE_W - MARGIN_LR * 2 }],
        children: [
          new TextRun({ text: sectionTitle || "" }),
          new TextRun({ text: "\t" }),
          new TextRun({ children: [PageNumber.CURRENT], style: "REF-PageNumber" }),
        ],
      }),
    ],
  });
}

function makeFooter() {
  return new Footer({ children: [new Paragraph({ children: [] })] });
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
  NAVY, INK, GRAY, LIGHT_GRAY, SUBTITLE_GRAY, CALLOUT_TEXT, PAPER,
  FONT_BODY, FONT_MONO, PAGE_W, PAGE_H, MARGIN_LR, MARGIN_TOP, MARGIN_BOTTOM,
  HEADER_H, FOOTER_H, CONTENT_W_IN, IMG_DIR,
  WORD_STYLES, DEFAULT_STYLES, CHARACTER_STYLES, CHAPTER_NUMBERING,
  imgDimsNative, imageRunFromFile, figureParagraph, figureRow, ruleBreak,
  caption, captionMulti, body, bodyDropCap, subheading, questionHeading, taskLabel,
  answerLabel, answerText, answerBox, calloutBox, chapterOpener, introHeading,
  formulaParagraph, formulaWhere,
  makeHeader, makeFooter, pageProps,
};
