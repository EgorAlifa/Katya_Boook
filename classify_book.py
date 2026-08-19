import json, re

d = json.load(open("parsed_raw.json", encoding="utf-8"))
parts = d["parts"]
tables = d["tables"]

def despace(s):
    return s.replace("\u2009", "").strip()

def plain(s):
    """Despace + strip markdown emphasis wrapper for robust pattern matching."""
    t = despace(s)
    t = re.sub(r"^\*+", "", t)
    t = re.sub(r"\*+$", "", t)
    return t.strip()

IMG_RE = re.compile(r'!\[[^\]]*\]\(media/(image\d+\.\w+)\)\{width="([\d.]+)in"\s*\n?\s*height="([\d.]+)in"\}')
CAP_RE = re.compile(r'\*\*(Рисунок[^*]+?)\*\*\s*\*([^*]*)\*')

def strip_md(s):
    s = s.replace("\u2009", "")
    s = re.sub(r"\*\*(.*?)\*\*", r"\1", s, flags=re.S)
    s = re.sub(r"\*(.*?)\*", r"\1", s, flags=re.S)
    s = re.sub(r"^>\s*", "", s, flags=re.M)
    s = re.sub(r"\s+", " ", s).strip()
    s = s.replace(" --- ", " — ").replace("--", "—")
    return s

out = []
i = 0
n = len(parts)
while i < n:
    p = parts[i]
    dp = despace(p)
    pl = plain(p)

    # chapter marker e.g. "ГЛАВА 1"
    m = re.match(r"^(ГЛАВА\s*\d+)$", pl)
    if m:
        out.append({"type": "chapter_mark", "text": m.group(1)})
        i += 1
        continue

    # ЗАКЛЮЧЕНИЕ marker
    if re.match(r"^ЗАКЛЮЧЕНИЕ$", pl):
        out.append({"type": "conclusion_mark"})
        i += 1
        continue

    # task label (bold or not, thin-spaced or not)
    m = re.match(r"^(ЗАДАЧА\s*\d+\s*-\s*\d+)$", pl)
    if m:
        out.append({"type": "task_label", "text": re.sub(r"\s*-\s*", "-", m.group(1))})
        i += 1
        continue

    # answer label
    if re.match(r"^ОТВЕТ$", pl):
        out.append({"type": "answer_label"})
        i += 1
        continue

    # pandoc "simple table" (dash-underlined columns) — used once for the
    # planet-symbol data table; capture as raw text, render as a plain table
    if re.match(r"^\s*-{15,}", p):
        out.append({"type": "dashtable", "raw": p})
        i += 1
        continue

    # table placeholder
    m = re.match(r"^@@TABLE(\d+)@@$", p.strip())
    if m:
        out.append({"type": "table", "raw": tables[int(m.group(1))]})
        i += 1
        continue

    # image-only block (possibly multiple images concatenated)
    imgs = IMG_RE.findall(p)
    if imgs and len(p) < 400 and p.strip().startswith("!["):
        out.append({"type": "image", "images": [{"file": f, "w": float(w), "h": float(h)} for f, w, h in imgs]})
        i += 1
        continue

    # figure caption(s) — possibly more than one caption in the same block
    caps = CAP_RE.findall(p)
    if caps and (dp.startswith("Рисунок") or p.strip().startswith("**Рисунок")):
        out.append({"type": "caption", "items": [{"num": strip_md(num), "desc": strip_md(desc)} for num, desc in caps]})
        i += 1
        continue
    # fallback caption variants: whole block wrapped uniformly in ** or * with no
    # separate bold/italic split (observed for a couple of irregular figures)
    m = re.match(r"^\*{1,2}Рисунок\s+([^*.]+)\.\s*(.*?)\*{1,2}$", p.strip(), flags=re.S)
    if m:
        out.append({"type": "caption", "items": [{"num": "Рисунок " + strip_md(m.group(1)) + ".", "desc": strip_md(m.group(2))}]})
        i += 1
        continue

    # subheading: blockquote + bold, OR standalone bold line under ~140 chars
    m = re.match(r"^>\s*\*\*(.+?)\*\*$", p.strip(), flags=re.S)
    if m and len(m.group(1)) < 140:
        out.append({"type": "subheading", "text": strip_md(m.group(1))})
        i += 1
        continue
    m = re.match(r"^\*\*(.+?)\*\*$", p.strip(), flags=re.S)
    if m and len(m.group(1)) < 140:
        txt = m.group(1)
        if despace(txt).isupper() and len(despace(txt)) > 4:
            out.append({"type": "chapter_title", "text": strip_md(txt)})
        else:
            out.append({"type": "subheading", "text": strip_md(txt)})
        i += 1
        continue

    # drop-cap paragraph: starts with **X** then lowercase continuation
    m = re.match(r"^\*\*([А-ЯA-Z])\*\*(.+)$", p, flags=re.S)
    if m:
        out.append({"type": "para_dropcap", "letter": m.group(1), "text": strip_md(m.group(2))})
        i += 1
        continue

    # blockquote — may contain an inline image glued to wrapped text (common
    # for task figures placed beside the question in the original layout)
    if p.strip().startswith(">"):
        dequote = re.sub(r"^>\s?", "", p, flags=re.M)
        inline_imgs = IMG_RE.findall(dequote)
        remainder = IMG_RE.sub("", dequote).strip()
        if inline_imgs:
            out.append({"type": "image", "images": [{"file": f, "w": float(w), "h": float(h)} for f, w, h in inline_imgs]})
        if remainder:
            mm = re.match(r"^\*\*(.+?)\*\*$", remainder, flags=re.S)
            if mm and len(mm.group(1)) < 140:
                out.append({"type": "subheading", "text": strip_md(mm.group(1))})
            else:
                out.append({"type": "para", "text": strip_md(remainder)})
        i += 1
        continue

    # default: regular paragraph (extract any inline image first — some
    # figures are glued mid-sentence in the original layout)
    inline_imgs = IMG_RE.findall(p)
    if inline_imgs:
        out.append({"type": "image", "images": [{"file": f, "w": float(w), "h": float(h)} for f, w, h in inline_imgs]})
        p = IMG_RE.sub("", p)
    text = strip_md(p)
    if text:
        out.append({"type": "para", "text": text})
    i += 1

json.dump(out, open("classified.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("blocks:", len(out))
from collections import Counter
print(Counter(b["type"] for b in out))

# ---- post-process grid-table blocks into structured callouts --------------
def parse_callout(raw):
    raw = raw.replace("\u2009", "")
    lines = raw.split("\n")
    cells = []
    for ln in lines:
        if ln.startswith("+"):
            continue
        # strip leading/trailing pipe & whitespace, split multi-column by "|"
        cols = [c.strip() for c in ln.strip("|").split("|")]
        cells.append(cols)
    ncols = max(len(c) for c in cells) if cells else 1
    # join column-wise (word-wrapped lines belonging to the same logical cell)
    coltext = ["" for _ in range(ncols)]
    for row in cells:
        for ci in range(ncols):
            val = row[ci] if ci < len(row) else ""
            if val:
                coltext[ci] += (" " if coltext[ci] else "") + val
    full = " ".join(t for t in coltext if t).strip()
    images = IMG_RE.findall(full.replace(" ", "")) or IMG_RE.findall(re.sub(r"\s+", "", full))
    # more robust image extraction: remove all whitespace first (wrapped mid-token)
    compact = re.sub(r"\s+", "", full)
    images = IMG_RE.findall(compact)
    textonly = IMG_RE.sub("", compact)
    # can't cleanly re-split compacted text back to words; instead extract images
    # from the ORIGINAL (non-compacted) joined text using a whitespace-tolerant regex
    img_re_loose = re.compile(r'!\[[^\]]*\]\(\s*media/\s*(image\d+)\s*\.\s*(\w+)\s*\)\{width="([\d.]+)in"\s*height="([\d.]+)in"\}')
    norm = re.sub(r"\s+", " ", full)
    norm2 = norm.replace("( media", "(media").replace(". ", ".")
    imgs2 = img_re_loose.findall(re.sub(r"\s*\n\s*", "", full))
    text_no_img = re.sub(r'!\[[^\]]*\]\([^)]*\)\{[^}]*\}', "", full)
    text_no_img = re.sub(r"\s+", " ", text_no_img).strip()

    dp = text_no_img
    body_text = dp
    body_text = re.sub(r"\*+", "", body_text)
    body_text = re.sub(r"^(ЭТО\s*ИНТЕРЕСНО|Это\s*интересно)\s*", "", body_text, flags=re.I).strip()
    return {"body_text": body_text, "images": imgs2}

for b in out:
    if b["type"] == "table":
        parsed = parse_callout(b["raw"])
        b["type"] = "callout"
        bt = parsed["body_text"]
        # detect an explicit bold subtitle inside the raw block (only the first
        # callout has one) — grab the second bold span, since the first is the
        # "ЭТО ИНТЕРЕСНО" label itself
        raw_clean = b["raw"].replace("\u2009", "")
        bolds = re.findall(r"\*\*(.+?)\*\*", raw_clean, flags=re.S)
        title = None
        if len(bolds) >= 2:
            cand = re.sub(r"\s+", " ", bolds[1]).strip()
            if 8 < len(cand) < 120 and cand in bt:
                title = cand
                bt = bt[len(title):].strip() if bt.startswith(title) else bt.replace(title, "", 1).strip()
        b["title"] = title
        b["body_text"] = bt
        b["images"] = [{"file": f + "." + ext, "w": float(w), "h": float(h)} for f, ext, w, h in parsed["images"]]

json.dump(out, open("classified.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("callouts parsed")
for b in out:
    if b["type"] == "callout":
        print("CALLOUT title:", b["title"], "| body:", b["body_text"][:80], "| imgs:", b["images"])
