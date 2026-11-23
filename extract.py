"""Arabic document extraction. Public entry points: extract() and write_outputs()."""
import argparse
import csv
import io
import json
import re
import unicodedata
from itertools import groupby
from pathlib import Path

import pdfplumber
import pytesseract
from docx import Document
from docx.table import Table
from PIL import Image, ImageSequence

IMAGE_TYPES = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".webp"}
SUPPORTED = IMAGE_TYPES | {".pdf", ".docx"}
MODES = ("auto", "always", "never")
OCR_THRESHOLD = 60
# Include punctuation inside addresses, URLs, dates, times and telephone numbers.
LTR_RUN = re.compile(r"[A-Za-z0-9\u0660-\u0669\u06f0-\u06f9@:.+/_#%=&?()\-]+"
                     r"(?: +[A-Za-z0-9\u0660-\u0669\u06f0-\u06f9@:.+/_#%=&?()\-]+)*")


def normalize(text):
    return unicodedata.normalize("NFKC", text)


def arabic(text):
    return any(unicodedata.bidirectional(c) == "AL" and c.isalpha() for c in text)


def record(page, kind, cells, review=False):
    return dict(page=page, kind=kind, review=review, cells=[normalize(s) for s in cells])


def lines(chars):
    groups = []
    for char in sorted(chars, key=lambda c: c["top"]):
        if not groups or char["top"] - groups[-1][0]["top"] > 3:
            groups.append([])
        groups[-1].append(char)
    return groups


def line_text(chars, rtl):
    ordered = sorted(chars, key=lambda c: c["x0"], reverse=rtl)
    pieces = []
    prev = None
    for char in ordered:
        if prev:
            gap = prev["x0"] - char["x1"] if rtl else char["x0"] - prev["x1"]
            if gap > 2 and not (prev["text"].isspace() or char["text"].isspace()):
                pieces.append(" ")
        pieces.append(char["text"])
        prev = char
    value = "".join(pieces).strip()
    if rtl:
        value = LTR_RUN.sub(lambda m: m[0][::-1], value)
    return normalize(value)


def inside(char, box):
    x0, top, x1, bottom = box
    return (char["x0"] >= x0 - .5 and char["x1"] <= x1 + .5
            and char["top"] >= top - .5 and char["bottom"] <= bottom + .5)


def pdf_rows(page, number):
    chars = page.dedupe_chars().chars
    events = []
    used = set()
    for table in page.find_tables():
        for index, row in enumerate(table.rows):
            nearby = [c for c in chars if inside(c, row.bbox)]
            cells = []
            for box in row.cells:
                selected = [c for c in nearby if box and inside(c, box)]
                used.update(id(c) for c in selected)
                rtl = any(arabic(c["text"]) for c in selected)
                cells.append("\n".join(line_text(line, rtl) for line in lines(selected)))
            if index == 0:
                rtl_table = any(arabic(value) for value in cells)
            events.append((row.bbox[1], record(number, "table", cells[::-1] if rtl_table else cells)))
    for line in lines(c for c in chars if id(c) not in used):
        value = line_text(line, any(arabic(c["text"]) for c in line))
        if value:
            events.append((line[0]["top"], record(number, "text", [value])))
    return [row for _, row in sorted(events, key=lambda e: e[0])]


def repair_ocr(text):
    words = text.split()
    if (not words or not re.search("[a-zA-Z]", words[0]) or not arabic(words[-1])
            or re.search(r"@|http|www", text, re.I)
            or sum(sum(arabic(c) for c in w) >= 2 for w in words) < 2):
        return text
    runs = [(rtl, " ".join(group)) for rtl, group in groupby(words, key=arabic)]
    if sum(not rtl and bool(re.search("[A-Za-z0-9]", s)) for rtl, s in runs) < 2:
        return text
    return " ".join(s for _, s in reversed(runs))


def ocr_rows(image, page):
    data = pytesseract.image_to_data(image, lang="ara+eng", config="--dpi 300 --psm 3",
                                    output_type=pytesseract.Output.DICT)
    grouped = {}
    for i, word in enumerate(data["text"]):
        word = normalize(word).strip()
        if word:
            key = tuple(data[k][i] for k in ("block_num", "par_num", "line_num"))
            grouped.setdefault(key, []).append((word, float(data["conf"][i])))
    return [record(page, "ocr", [repair_ocr(" ".join(w for w, _ in line))],
                   min(conf for _, conf in line) < OCR_THRESHOLD) for line in grouped.values()]


def extract(path, ocr="auto"):
    path = Path(path)
    if ocr not in MODES or path.suffix.lower() not in SUPPORTED:
        raise ValueError("نوع المستند أو وضع OCR غير مدعوم.")
    rows = []
    if path.suffix.lower() == ".pdf":
        with pdfplumber.open(path) as doc:
            for number, page in enumerate(doc.pages, 1):
                if any(c["text"].strip() for c in page.chars):
                    rows.extend(pdf_rows(page, number))
                elif ocr != "never":
                    with page.to_image(resolution=300).original as image:
                        rows.extend(ocr_rows(image, number))
                page.close()
    elif path.suffix.lower() == ".docx":
        for block in Document(path).iter_inner_content():
            if isinstance(block, Table):
                rows.extend(record(None, "table", [c.text for c in r.cells]) for r in block.rows)
            else:
                rows.extend(record(None, "text", [s]) for s in block.text.splitlines() if s.strip())
    elif ocr != "never":
        with Image.open(path) as image:
            for number, frame in enumerate(ImageSequence.Iterator(image), 1):
                rows.extend(ocr_rows(frame, number))
    return rows


def flat_rows(rows):
    width = max((len(r["cells"]) for r in rows), default=1)
    yield ["page", "kind", "review"] + [f"cell_{i+1}" for i in range(width)]
    for r in rows:
        yield [r["page"], r["kind"], int(r["review"]), *r["cells"], *([""] * (width-len(r["cells"])))]


def exports(rows):
    stream = io.StringIO(newline="")
    csv.writer(stream).writerows(flat_rows(rows))
    return {"csv": stream.getvalue().encode("utf-8"),
            "json": json.dumps(rows, ensure_ascii=False, indent=2).encode("utf-8")}


def write_outputs(rows, outdir):
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    for ext, data in exports(rows).items():
        (outdir / f"result.{ext}").write_bytes(data)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="استخراج صفوف المستند")
    parser.add_argument("input", type=Path)
    parser.add_argument("-o", required=True, type=Path, metavar="OUTDIR")
    parser.add_argument("--ocr", choices=MODES, default="auto")
    args = parser.parse_args()
    try:
        write_outputs(extract(args.input, args.ocr), args.o)
    except Exception as error:
        parser.exit(1, f"{error}\n")
