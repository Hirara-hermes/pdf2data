import base64
import csv
import io
import json
import subprocess
import sys
import tempfile
import unittest
from html.parser import HTMLParser
from pathlib import Path
from unittest.mock import patch

from docx import Document
from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont

from app import app
from extract import extract, normalize, repair_ocr

FONT = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
# Each pair holds independently specified visual glyph order and logical text.
NAME = ("ﻢﺳﺍ", "اسم")
VALUE = ("ﺔﻤﻴﻗ", "قيمة")
AR_HEADER = [NAME, VALUE]


def same(text):
    return (text, text)


def make_pdf(path, logical_rows, rtl):
    try:
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        from reportlab.pdfgen import canvas
    except ImportError:
        from fpdf import FPDF
        pdf = FPDF(unit="pt", format=(550, 400))
        pdf.add_page()
        pdf.add_font("TestArabic", fname=FONT)
        pdf.set_font("TestArabic", size=12)
        line, label = pdf.line, pdf.text
        save = lambda: pdf.output(path)
    else:
        pdfmetrics.registerFont(TTFont("TestArabic", FONT))
        pdf = canvas.Canvas(str(path), pagesize=(550, 400))
        pdf.setFont("TestArabic", 12)

        def line(x, y, end_x, end_y):
            pdf.line(x, 400-y, end_x, 400-end_y)

        def label(x, y, text):
            pdf.drawString(x, 400-y, text)

        save = pdf.save

    label(20, 20, "Before")
    for x in (20, 270, 520):
        line(x, 40, x, 40 + 30 * len(logical_rows))
    for i in range(len(logical_rows) + 1):
        line(20, 40 + 30*i, 520, 40 + 30*i)
    for i, row in enumerate(logical_rows):
        for j, (visual, _) in enumerate(reversed(row) if rtl else row):
            label(25 + 250*j, 60 + 30*i, visual)
    label(20, 70 + 30*len(logical_rows), "After")
    save()


class Links(HTMLParser):
    def __init__(self):
        super().__init__()
        self.files = {}

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if tag == "a" and "download" in attrs:
            self.files[attrs["download"]] = base64.b64decode(attrs["href"].split(",", 1)[1])


class ExtractionTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory(dir=Path(__file__).resolve().parents[1])
        self.addCleanup(temp.cleanup)
        self.folder = Path(temp.name)

    def test_pdf_tables(self):
        cases = [
            (AR_HEADER, [NAME, VALUE]),
            ([same("Name"), same("Value")], [NAME, VALUE]),
            (AR_HEADER, [same("AWS"), same("PMP")]),
            (AR_HEADER, [NAME, same("")]),
            (AR_HEADER, [("١٢:٣٠ ﺔﻤﻴﻗ", "قيمة ١٢:٣٠"), same("2026-11-23")]),
            (AR_HEADER, [("a.b@example.com ﻢﺳﺍ", "اسم a.b@example.com"), same("AWS-SAA")]),
            (AR_HEADER, [("+966 55 123 4567 ﺔﻤﻴﻗ", "قيمة +966 55 123 4567"), same("3000.50")]),
            (AR_HEADER, [("https://abc.net/a?x=2 ﻢﺳﺍ", "اسم https://abc.net/a?x=2"),
                         same("Acme Inc.")]),
        ]
        for i, (header, data) in enumerate(cases):
            with self.subTest(i=i):
                path = self.folder / f"{i}.pdf"
                make_pdf(path, [header, data, data], header == AR_HEADER)
                truth = [["Before"], *[[c[1] for c in r] for r in [header, data, data]], ["After"]]
                truth_path = path.with_suffix(".truth.json")
                truth_path.write_text(json.dumps(truth, ensure_ascii=False), encoding="utf-8")
                with patch("extract.ocr_rows", side_effect=AssertionError("OCR on text")):
                    for mode in ("auto", "always", "never"):
                        rows = extract(path, mode)
                        self.assertEqual([r["cells"] for r in rows], json.loads(truth_path.read_bytes()))
                        self.assertEqual([r["kind"] for r in rows], ["text", "table", "table", "table", "text"])

    def test_nfkc_and_cli(self):
        self.assertEqual(normalize("اﻟﺴﻴﺮة"), "السيرة")
        doc = Document()
        doc.add_paragraph("اﻟﺴﻴﺮة")
        doc.add_table(1, 1).cell(0, 0).text = "ＡＷＳ"
        path = self.folder / "input.docx"
        doc.save(path)
        subprocess.run([sys.executable, "extract.py", path, "-o", self.folder], check=True)
        self.assertEqual([r["cells"] for r in json.loads((self.folder/"result.json").read_bytes())],
                         [["السيرة"], ["AWS"]])
        with (self.folder/"result.csv").open(encoding="utf-8", newline="") as stream:
            self.assertEqual([r["cell_1"] for r in csv.DictReader(stream)], ["السيرة", "AWS"])

    def test_web_downloads(self):
        doc = Document()
        doc.add_paragraph("=1+2")
        stream = io.BytesIO()
        doc.save(stream)
        stream.seek(0)
        client = app.test_client()
        self.assertEqual(client.get("/").status_code, 200)
        response = client.post("/extract", data={"file": (stream, "جدول.docx")})
        self.assertEqual(response.status_code, 200)
        links = Links()
        links.feed(response.get_data(as_text=True))
        self.assertEqual(set(links.files), {"result.csv", "result.json", "result.xlsx"})
        cell = load_workbook(io.BytesIO(links.files["result.xlsx"])).active["D2"]
        self.assertEqual((cell.value, cell.data_type), ("=1+2", "s"))
        self.assertEqual(client.post("/extract").status_code, 400)
        with patch.dict(app.config, MAX_CONTENT_LENGTH=100):
            response = client.post("/extract", data={"file": (io.BytesIO(b"x"*101), "x.pdf")})
            self.assertEqual(response.status_code, 413)

    def test_scan_and_image(self):
        image = Image.new("RGB", (1200, 500), "white")
        draw = ImageDraw.Draw(image)
        for i, text in enumerate(["Invoice total", "Total 456"]):
            draw.text((100, 100 + 100*i), text, font=ImageFont.truetype(FONT, 56), fill="black")
        for ext in ("png", "pdf"):
            path = self.folder / ("scan." + ext)
            image.save(path, dpi=(300, 300))
            self.assertEqual(extract(path, "never"), [])
            rows = extract(path)
            self.assertEqual([r["cells"] for r in rows], [["Invoice total"], ["Total 456"]])
            self.assertTrue(all(r["kind"] == "ocr" for r in rows))

    def test_mixed_scan_error_rates(self):
        def distance(a, b):
            previous = list(range(len(b) + 1))
            for i, left in enumerate(a, 1):
                current = [i]
                for j, right in enumerate(b, 1):
                    current.append(min(current[-1] + 1, previous[j] + 1,
                                       previous[j-1] + (left != right)))
                previous = current
            return previous[-1]

        truth = ["Invoice 123", "السيرة الذاتية", "Total 456"]
        image = Image.new("RGB", (1600, 650), "white")
        font = ImageFont.truetype(FONT, 60)
        for i, value in enumerate(truth):
            ImageDraw.Draw(image).text((100, 100+i*140), value, font=font, fill="black")
        path = self.folder / "mixed.png"
        image.save(path, dpi=(300, 300))
        actual = [row["cells"][0] for row in extract(path)]
        self.assertEqual(len(actual), len(truth))
        matches = [min(actual, key=lambda line: distance(ref, line)) for ref in truth]
        self.assertCountEqual(matches, actual)
        for tokenize, limit in ((list, 0.15), (str.split, 0.35)):
            errors = sum(distance(tokenize(a), tokenize(b)) for a, b in zip(truth, matches))
            self.assertLessEqual(errors / sum(len(tokenize(s)) for s in truth), limit)

    def test_ocr_segments(self):
        for s in ("Name: أحمد علي", "+966 55 123 4567", "AWS أحمد a@x.net علي",
                  "AWS أحمد https://x.net علي", "AWS أحمد www.x.net علي", "علي AWS أحمد PMP"):
            self.assertEqual(repair_ocr(s), s)
        self.assertEqual(repair_ocr("AWS محمد علي PMP أحمد حسن"), "أحمد حسن PMP محمد علي AWS")
