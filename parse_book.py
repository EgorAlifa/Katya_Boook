import re, json, sys

TXT = open("book.md", encoding="utf-8").read()

def despace(s):
    """Remove pandoc's thin-space (U+2009) letter spacing used for small-caps-style labels."""
    return s.replace("\u2009", "")

def is_spaced_caps(s, plain_regex):
    d = despace(s).strip()
    return re.match(plain_regex, d)

# ---- locate main body range -------------------------------------------------
start_idx = TXT.index("**В\u2009В\u2009Е\u2009Д\u2009Е\u2009Н\u2009И\u2009Е**") if "\u2009" in TXT[:60] else TXT.index("**В")
# more robust: find literal intro heading token
m = re.search(r"\*\*В\u2009?В\u2009?Е\u2009?Д\u2009?Е\u2009?Н\u2009?И\u2009?Е\*\*", TXT)
start_idx = m.start()
m2 = re.search(r"\*\*С\u2009?О\u2009?Д\u2009?Е\u2009?Р\u2009?Ж\u2009?А\u2009?Н\u2009?И\u2009?Е\*\*", TXT)
toc_idx = m2.start()

body = TXT[start_idx:toc_idx]
tail = TXT[toc_idx:]

# ---- split into blank-line separated blocks, but keep grid tables intact ---
raw_lines = body.split("\n")

blocks = []
buf = []
in_table = False
for line in raw_lines:
    if line.startswith("+---") or line.startswith("+==="):
        in_table = True
        buf.append(line)
        continue
    if in_table:
        buf.append(line)
        if line.startswith("+---") and len(buf) > 1:
            # could be closing border of the table (heuristic: next non-table line is blank)
            pass
        continue
    if line.strip() == "":
        if buf:
            blocks.append("\n".join(buf))
            buf = []
    else:
        buf.append(line)
if buf:
    blocks.append("\n".join(buf))

# second pass: tables were being swallowed line-by-line without proper closing detection;
# redo with a cleaner state machine
blocks = []
buf = []
mode = "normal"
for line in raw_lines:
    if mode == "normal":
        if line.startswith("+---"):
            mode = "table"
            buf = [line]
            continue
        if line.strip() == "":
            if buf:
                blocks.append(("text", "\n".join(buf)))
                buf = []
        else:
            buf.append(line)
    elif mode == "table":
        buf.append(line)
        if line.startswith("+---") and not line.startswith("+==="):
            # count border lines; a table ends with a +--- line followed by blank line
            pass
        # detect end: this +--- line, then peek not available; use heuristic below after loop
if buf:
    blocks.append(("text", "\n".join(buf)))

# The state machine above is awkward for detecting table end inline; instead use regex to
# pull out grid tables first, replacing them with placeholders, then split the rest normally.
TABLE_RE = re.compile(r"(\+-{5,}.*?\+-{5,}\+)\n\n", re.S)

def extract_tables(text):
    tables = []
    def repl(m):
        tables.append(m.group(1))
        return f"\n\n@@TABLE{len(tables)-1}@@\n\n"
    new_text = TABLE_RE.sub(repl, text + "\n\n")
    return new_text, tables

body2, tables = extract_tables(body)
parts = [p.strip("\n") for p in re.split(r"\n\s*\n", body2) if p.strip()]

print(f"Found {len(tables)} grid tables", file=sys.stderr)
print(f"Found {len(parts)} top-level blocks", file=sys.stderr)

json.dump({"parts": parts, "tables": tables, "tail": tail}, open("parsed_raw.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("wrote parsed_raw.json")
