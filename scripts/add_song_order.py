#!/usr/bin/env python3
"""
Add the performance order (from the PDF chord/lyric charts) to your lyric .txt files.

For each PDF in a folder it finds the .txt with the same base filename, then:
  1. parses the .txt into labelled sections (Verse 1, Chorus, Tag, ...)
  2. reads the PDF text and walks through it in order:
     - section headers are detected with the same regexes as normalizeLyricStructure
       in src/presenter.js:62-82 ([Verse 1], (Chorus), Verse 1:, Chorus:, ...)
     - inline order / repeat notes are expanded: Chorus x2, Bridge (x3), (repeat),
       (REPEAT), x2, Repeat verse 2, bare Chorus / Bridge / Tag / Verse 1, ...
     - lyric blocks between markers are matched to .txt sections by content
  3. writes the order back to the .txt

Usage:
  python add_song_order.py FOLDER               # preview only (dry run)
  python add_song_order.py FOLDER --write       # add an "Order:" line to the header
  python add_song_order.py FOLDER --write --expand
        # additionally rewrite the body so sections appear in performance order,
        # repeating the lyrics under each header (for apps that play sections top to bottom)

A .txt.bak backup is written before any file is modified. Review the dry-run
output first: PDF text extraction and section naming are never 100% reliable,
so some songs will need the Order set manually or the .txt/.pdf aligned.

Install one of:  pip install pymupdf   (preferred)   or   pip install pdfplumber
"""
import argparse
import re
import sys
from pathlib import Path

# ---------------------------------------------------------------------------
# Section header detection - ported from normalizeLyricStructure
# (src/presenter.js:62-82)
# ---------------------------------------------------------------------------
SECTION_KEYS = (
    'verse', 'chorus', 'bridge', 'pre-chorus', 'prechorus', 'tag', 'outro',
    'intro', 'ending', 'refrain', 'hook', 'v', 'c', 'b',
)
_SECTION = '|'.join(SECTION_KEYS)
# [Verse 1] / (Chorus) / [v2]   (bracket/paren regex at presenter.js:67)
BRACKET_RE = re.compile(
    r'^\s*[\[\(]\s*(' + _SECTION + r')\s*(\d*[a-z]?)\s*[\]\)]\s*:?\s*$', re.I)
# Verse 1: / Chorus: / Pre-chorus:  (colon regex at presenter.js:74)
COLON_RE = re.compile(
    r'^\s*(' + _SECTION.replace('|v|c|b', '') + r')\s*(\d*[a-z]?)\s*:\s*$', re.I)
# Bare headers as stored in .txt files: Verse 1 / Chorus / Tag / Pre-Chorus
BARE_RE = re.compile(
    r'^\s*(' + _SECTION.replace('|v|c|b', '') + r')\s*(\d*[a-z]?)\s*$', re.I)

_LABEL_MAP = {
    'v': 'Verse', 'c': 'Chorus', 'b': 'Bridge',
    'prechorus': 'Pre-Chorus', 'pre-chorus': 'Pre-Chorus',
    'ending': 'Outro', 'refrain': 'Chorus', 'hook': 'Chorus',
}


def norm_text(s: str) -> str:
    s = s.lower().replace("\u2019", "'").replace("\u2018", "'").replace("\ufffd", "")
    return re.sub(r"[^a-z0-9']+", " ", s).strip()


def norm_label(key: str, num: str) -> str:
    key = key.lower().replace(' ', '')
    name = _LABEL_MAP.get(key, key[:1].upper() + key[1:].lower())
    if num:
        name = name + ' ' + num
    return name


def detect_header(line: str):
    """Return normalized section label if line is a section header, else None."""
    for rx in (BRACKET_RE, COLON_RE, BARE_RE):
        m = rx.match(line)
        if m:
            return norm_label(m.group(1), m.group(2))
    return None


# ---------------------------------------------------------------------------
# PDF order / repeat markers
# ---------------------------------------------------------------------------
# Chorus / Bridge x2 / Bridge (x2) / Repeat verse 2 / Chorus 1 / Pre-chorus
SECTION_MARKER_RE = re.compile(
    r'^\(?\s*(?:repeat\s+)?(' + _SECTION + r')\s*(\d*[a-z]?)?\s*'
    r'(?:\(?\s*x\s*(\d+)\s*\)?)?\s*\)?$', re.I)
# (repeat) / (REPEAT) / x2 / x3 / (x 2) / 2x
BARE_REPEAT_RE = re.compile(
    r'^\(?\s*(?:repeat|repeats?)\s*\)?$'
    r'|^\(\s*x\s*(\d+)\s*\)$|^x(\d+)$|^(\d+)x$', re.I)
REPEAT_FROM_TOP_RE = re.compile(r'^\(\s*repeat\s+from\s+top\s*\)$', re.I)

META_PREFIXES = ("ccli", "©", "(c)", "copyright", "key", "tempo", "capo", "bpm", "time")


def classify(line: str):
    """Classify a PDF line: ('meta',None) | ('header',label) | ('marker',(label,times))
    | ('repeat',times) | ('repeat_top',None) | ('lyric',text)."""
    low = line.lower()
    if low.startswith(META_PREFIXES) or ' | ' in low:
        return ('meta', None)
    if re.search(r'\b(written by|lyrics by|music by|arranged by|source:)\b', low):
        return ('meta', None)
    h = detect_header(line)
    if h:
        return ('header', h)
    m = SECTION_MARKER_RE.match(line)
    if m:
        return ('marker', (norm_label(m.group(1), m.group(2)), int(m.group(3) or 1)))
    m = REPEAT_FROM_TOP_RE.match(line)
    if m:
        return ('repeat_top', None)
    m = BARE_REPEAT_RE.match(line)
    if m:
        times = int(m.group(1) or m.group(2) or m.group(3) or 1)
        return ('repeat', times)
    return ('lyric', line)


# ---------------------------------------------------------------------------
# .txt parsing
# ---------------------------------------------------------------------------
def split_txt(path: Path):
    """Split a .txt into header (meta lines before the first section) and
    [(label, [body lines])] sections. Section headers are detected with the
    same normalizeLyricStructure regexes used above."""
    header, sections, cur = [], [], None
    for raw in path.read_text(encoding='utf-8').splitlines():
        line = raw.rstrip()
        lab = detect_header(line.strip())
        if lab:
            cur = (lab, [])
            sections.append(cur)
        elif cur is None:
            header.append(line)
        elif line.strip():
            cur[1].append(line.strip())
    return header, sections


def canon_label(label: str) -> str:
    return norm_text(label).replace('pre chorus', 'prechorus')


# ---------------------------------------------------------------------------
# Build performance order from PDF text
# ---------------------------------------------------------------------------
def pdf_order(pdf_lines, sections):
    """
    Walk the PDF in order and produce [(label, body_lines), ...] pairs.
    Section labels come from PDF headers/markers; lyric bodies come from
    matching the PDF lyric blocks to .txt section content (longest-prefix
    match, skipping unmatched leading lines such as song titles).
    """
    bodies_orig = {lab: list(body) for lab, body in sections}
    bodies_norm = {lab: [norm_text(l) for l in body] for lab, body in sections}
    by_label = {canon_label(lab): lab for lab, _ in sections}

    def resolve(label):
        key = canon_label(label)
        if key in by_label:
            return by_label[key]
        for k, lab in by_label.items():
            if k.startswith(key) or key.startswith(k):
                return lab
        return None

    def match_section(nl, start=0):
        """Match a .txt section body against nl starting at position `start`.
        Tries every position >= start where a section's first line matches.
        Returns (label, skip_count, matched_len) or None."""
        if not nl or start >= len(nl):
            return None
        first_lines = {sec[0] for sec in bodies_norm.values() if sec}
        best = None
        best_skip = 0
        best_len = 0
        for skip in range(start, len(nl)):
            if nl[skip] not in first_lines:
                continue
            for lab, sec_norm in bodies_norm.items():
                if not sec_norm:
                    continue
                if len(sec_norm) <= len(nl) - skip and nl[skip:skip + len(sec_norm)] == sec_norm:
                    if len(sec_norm) > best_len:
                        best = lab
                        best_skip = skip
                        best_len = len(sec_norm)
        if best:
            return best, best_skip, best_len
        return None

    order = []        # list of (label, body_lines)
    warnings = []
    last = None       # last emitted section label
    pending = None    # (label, count) from a header/marker awaiting its lyric block

    def flush():
        nonlocal last, pending
        if pending:
            plabel, pcount = pending
            resolved = resolve(plabel)
            if resolved:
                order.extend([(resolved, bodies_orig[resolved])] * pcount)
                last = resolved
            else:
                warnings.append(f"Section '{plabel}' not found in .txt")
            pending = None

    i = 0
    while i < len(pdf_lines):
        kind, payload = classify(pdf_lines[i])
        if kind == 'meta':
            i += 1
            continue
        if kind == 'header':
            flush()
            pending = (payload, 1)
            i += 1
            continue
        if kind == 'marker':
            flush()
            label, times = payload
            resolved = resolve(label)
            if resolved:
                order.extend([(resolved, bodies_orig[resolved])] * times)
                last = resolved
            else:
                warnings.append(f"Section '{label}' not found in .txt")
            pending = None
            i += 1
            continue
        if kind == 'repeat_top':
            flush()
            snapshot = order.copy()
            order.extend(snapshot)
            warnings.append("(Repeat from top): repeated all preceding sections")
            i += 1
            continue
        if kind == 'repeat':
            flush()
            if last:
                order.extend([(last, bodies_orig[last])] * payload)
            else:
                warnings.append("Repeat marker with no preceding section to repeat")
            i += 1
            continue
        # Lyric line: accumulate into the current block
        block = [pdf_lines[i]]
        i += 1
        while i < len(pdf_lines):
            k2, _ = classify(pdf_lines[i])
            if k2 != 'lyric':
                break
            block.append(pdf_lines[i])
            i += 1
        flush()  # emit any pending header/marker before matching this block
        # Match and split the block into one or more sections
        nl = [norm_text(l) for l in block]
        pos = 0
        while pos < len(nl):
            result = match_section(nl, pos)
            if result:
                lab, skip, mlen = result
                if skip > pos:
                    warnings.append(f"Unmatched lyric text: {nl[pos][:60]!r}")
                order.append((lab, bodies_orig[lab]))
                last = lab
                pos = skip + mlen
            else:
                warnings.append(f"Unmatched lyric text: {nl[pos][:60]!r}")
                pos += 1

    flush()
    return order, warnings


# ---------------------------------------------------------------------------
# Output writing
# ---------------------------------------------------------------------------
def build_expanded(header, order):
    out = list(header)
    while out and not out[-1].strip():
        out.pop()
    for lab, body in order:
        out += ['', lab, *body]
    return '\n'.join(out) + '\n'


def add_order_line(header, order_labels):
    line = 'Order: ' + ', '.join(order_labels)
    header = [h for h in header if not h.lower().startswith('order:')]
    while header and not header[-1].strip():
        header.pop()
    return header + [line]


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("folder")
    ap.add_argument("--write", action="store_true", help="actually modify the .txt files")
    ap.add_argument("--expand", action="store_true", help="rewrite body in performance order")
    ap.add_argument("--output", help="output folder for finalized .txt files (default: alongside originals)")
    args = ap.parse_args()

    folder = Path(args.folder)
    if not folder.is_dir():
        sys.exit(f"Folder not found: {folder}")

    out_folder = Path(args.output) if args.output else folder
    if args.output:
        out_folder.mkdir(parents=True, exist_ok=True)

    try:
        import pymupdf as fitz
    except ImportError:
        try:
            import fitz
        except ImportError:
            fitz = None

    def pdf_lines(path):
        if fitz is not None:
            with fitz.open(path) as doc:
                text = "\n".join(page.get_text() for page in doc)
        else:
            try:
                import pdfplumber
            except ImportError:
                sys.exit("Install pymupdf or pdfplumber:  pip install pymupdf")
            with pdfplumber.open(path) as pdf:
                text = "\n".join((p.extract_text() or "") for p in pdf.pages)
        lines = [l.strip() for l in text.splitlines() if l.strip()]
        while lines:
            kind, _ = classify(lines[0])
            if kind == 'meta':
                lines.pop(0)
            else:
                break
        return lines

    total = written = skipped = 0
    processed_txts = set()
    for pdf in sorted(folder.glob("*.pdf")):
        txt = pdf.with_suffix(".txt")
        processed_txts.add(txt.name.lower())
        if not txt.exists():
            print(f"SKIP  {pdf.name}: no matching .txt")
            skipped += 1
            continue
        total += 1
        header, sections = split_txt(txt)
        order, warns = pdf_order(pdf_lines(pdf), sections)
        order_labels = [lab for lab, _ in order]
        used = set(order_labels)
        missing = [lab for lab, _ in sections if lab not in used]

        # Fall back to original .txt section order for any missing sections
        final_order = list(order)
        for lab, body in sections:
            if lab not in used:
                final_order.append((lab, body))
                missing.remove(lab)
        order_labels = [lab for lab, _ in final_order]

        print(f"\n{txt.name}")
        print(f"  Order: {', '.join([lab for lab, _ in final_order]) or '(none found)'}")
        for w in warns:
            print("  WARN:", w)
        if missing:
            print(f"  NOTE: sections not found in PDF: {missing}  (using original .txt order)")

        if args.write:
            header_out = add_order_line(header, order_labels)
            if args.expand:
                text = build_expanded(header_out, final_order)
            else:
                text = "\n".join(header_out) + "\n"
                for lab, body in sections:
                    text += "\n" + lab + "\n" + "\n".join(body) + "\n"
            out_path = out_folder / txt.name
            bak = out_path.with_suffix(".txt.bak")
            if out_path.exists():
                bak.write_text(out_path.read_text(encoding="utf-8"), encoding="utf-8")
            out_path.write_text(text, encoding="utf-8")
            written += 1
            print(f"  written -> {out_path}")

    # Also process .txt files with no matching PDF (keep original order, add Order line)
    for txt in sorted(folder.glob("*.txt")):
        if txt.name.lower() in processed_txts:
            continue
        total += 1
        header, sections = split_txt(txt)
        final_order = list(sections)
        order_labels = [lab for lab, _ in final_order]

        print(f"\n{txt.name}")
        print(f"  Order: {', '.join(order_labels) or '(none found)'}  (no PDF - original order)")

        if args.write:
            header_out = add_order_line(header, order_labels)
            if args.expand:
                text = build_expanded(header_out, final_order)
            else:
                text = "\n".join(header_out) + "\n"
                for lab, body in sections:
                    text += "\n" + lab + "\n" + "\n".join(body) + "\n"
            out_path = out_folder / txt.name
            bak = out_path.with_suffix(".txt.bak")
            if out_path.exists():
                bak.write_text(out_path.read_text(encoding="utf-8"), encoding="utf-8")
            out_path.write_text(text, encoding="utf-8")
            written += 1
            print(f"  written -> {out_path}")

    print(f"\n{total} pairs checked, {written} written, {skipped} skipped")


if __name__ == "__main__":
    main()
