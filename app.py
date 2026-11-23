import os
import base64
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory

from flask import Flask, render_template_string, request
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from pytesseract import TesseractError, TesseractNotFoundError

from extract import SUPPORTED, exports, extract, flat_rows

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024

PAGE = """<!doctype html><html lang="ar" dir="rtl"><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>استخراج المستندات</title>
<style>body{font:1.1rem system-ui;max-width:70rem;margin:3rem auto;padding:0 1rem;color:#163337}
form,nav{display:flex;gap:1rem;align-items:center;flex-wrap:wrap}
table{border-collapse:collapse;margin-top:1rem}td,th{padding:.5rem;border:1px solid #aaa;white-space:pre-wrap}
button{padding:.45rem 1rem}a{color:#075c63}</style>
<h1>استخراج المستندات</h1>
<form action="/extract" method="post" enctype="multipart/form-data">
<label>اختر PDF أو DOCX أو صورة (حتى 25 ميبيبايت)
<input required type="file" name="file" accept="{{ accept }}"></label><button>استخراج</button></form>
{% if error %}<p role="alert">{{ error }}</p>{% endif %}
{% if rows is defined %}<h2>النتيجة</h2><nav aria-label="تنزيل النتائج">
{% for ext,uri in links.items() %}<a download="result.{{ ext }}" href="{{ uri }}">تنزيل {{ ext|upper }}</a>{% endfor %}
</nav>{% if rows %}<table><thead><tr><th>الصفحة</th><th>المراجعة</th>
{% for i in range(width) %}<th>حقل {{ i+1 }}</th>{% endfor %}</tr></thead><tbody>
{% for row in rows %}<tr><td>{{ row.page or '—' }}</td><td>{{ 'ثقة منخفضة' if row.review else '' }}</td>
{% for i in range(width) %}<td dir="auto">{{ row.cells[i] }}</td>{% endfor %}</tr>{% endfor %}
</tbody></table>{% else %}<p>لم تُستخرج صفوف.</p>{% endif %}{% endif %}</html>"""


def page(**context):
    return render_template_string(PAGE, accept=",".join(sorted(SUPPORTED)), **context)


@app.get("/")
def index():
    return page()


@app.errorhandler(413)
def too_large(_error):
    return page(error="تجاوز حجم الطلب الحد المسموح (25 ميبيبايت)."), 413


def downloads(rows):
    files = exports(rows)
    book = Workbook(write_only=True)
    sheet = book.create_sheet("النتائج")
    sheet.sheet_view.rightToLeft = True
    for row in flat_rows(rows):
        cells = [WriteOnlyCell(sheet, value=value) for value in row]
        for cell in cells:
            if isinstance(cell.value, str):
                cell.data_type = "s"
        sheet.append(cells)
    buffer = BytesIO()
    book.save(buffer)
    files["xlsx"] = buffer.getvalue()
    types = {"json": "application/json", "csv": "text/csv;charset=utf-8",
             "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}
    return {ext: f"data:{types[ext]};base64,{base64.b64encode(data).decode('ascii')}"
            for ext, data in files.items()}


@app.post("/extract")
def upload():
    file = request.files.get("file")
    suffix = Path(file.filename or "").suffix.lower() if file else ""
    if suffix not in SUPPORTED:
        return page(error="اختر مستندًا أو صورة من الأنواع المدعومة."), 400
    with TemporaryDirectory(dir=Path(__file__).resolve().parent) as folder:
        path = Path(folder) / ("upload" + suffix)
        file.save(path)
        try:
            rows = extract(path)
        except (TesseractError, TesseractNotFoundError):
            return page(error="تعذر تشغيل OCR. تحقق من تثبيت Tesseract ولغتي العربية والإنجليزية."), 503
        except Exception:
            app.logger.exception("Document extraction failed")
            return page(error="تعذرت قراءة المستند. تحقق من سلامة الملف ونوعه."), 422
    return page(rows=rows, links=downloads(rows), width=max((len(r["cells"]) for r in rows), default=0))


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=int(os.environ.get("PORT", "5061")))
