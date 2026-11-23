# استخراج المستندات العربية

يتطلب Python 3.10+ وTesseract مع حزمتَي اللغة `ara` و`eng`. من مجلد المشروع شغّل:

```sh
python -m venv .venv
. .venv/bin/activate
python -m pip install 'pdfplumber>=0.11' pytesseract Pillow 'python-docx>=1.1' 'Flask>=3' openpyxl reportlab pytest
python extract.py input.pdf -o out --ocr auto
python app.py
python -m pytest -q
```

تفتح واجهة الرفع على http://127.0.0.1:5000 وتقبل طلبات حتى 25 ميبيبايت. تُحذف نسخة الرفع بعد معالجتها، وتُضمّن روابط CSV وJSON وXLSX في صفحة النتيجة نفسها.

تكتب واجهة الأوامر `result.json` و`result.csv` بترميز UTF-8. تعيد الدالة `extract(path, ocr="auto")` قائمة صفوف، ولكل صف رقم صفحة (فارغ في DOCX)، ونوع، وعلامة مراجعة، وخلايا مطبّعة بصيغة NFKC. يتبع ترتيب أعمدة جدول PDF لغة الترويسة.

يعمل `auto` و`always` على الصور وصفحات PDF الخالية من النص فقط. يستخدم كلاهما Tesseract بالعربية والإنجليزية، وتُصيَّر صفحات PDF الممسوحة عند 300 dpi. يعطّل `never` التعرف الضوئي ويتخطى تلك الصفحات. تُعرض نتائج OCR كأسطر، ويُوسم السطر للمراجعة إذا انخفضت ثقة أي كلمة فيه عن 60 من 100.
