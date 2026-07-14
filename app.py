import base64
import io
import re
import zipfile
from flask import Flask, request, Response

from docx import Document

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # до 200 МБ на запит


# ---------- Нормалізація назв ----------

def normalize_name(s):
    s = (s or '')
    s = s.replace('_', ' ').replace('\u00A0', ' ')
    s = re.sub(r'\s+', ' ', s).strip()
    return s.lower()


# ---------- Побудова регулярки й резолвера з мапи даних ----------

def build_matcher(data_map):
    # data_map: {original_column_name: value}
    lookup = {}
    names = []
    for key, val in data_map.items():
        norm = normalize_name(key)
        if norm:
            lookup[norm] = '' if val is None else str(val)
            names.append(key)

    if not names:
        return None, {}

    # довші назви — раніше, щоб уникнути часткових збігів
    names.sort(key=lambda n: len(n), reverse=True)

    alts = []
    for n in names:
        piece = re.escape(n.strip())
        # пробіли в назві = пробіл або підкреслення (одне чи кілька)
        piece = re.sub(r'\\?\s+', r'[\\s_]+', piece)
        alts.append(piece)

    pattern = re.compile(
        '\u00AB[\\s_]*(' + '|'.join(alts) + ')[\\s_]*\u00BB',
        re.IGNORECASE | re.UNICODE
    )
    return pattern, lookup


# ---------- Заміна в одному абзаці зі збереженням форматування ----------

def replace_in_paragraph(paragraph, pattern, lookup):
    runs = paragraph.runs
    if not runs:
        return
    if '\u00AB' not in ''.join(r.text for r in runs):
        return

    guard = 0
    while True:
        guard += 1
        if guard > 2000:
            break

        texts = [r.text for r in runs]
        full = ''.join(texts)
        m = pattern.search(full)
        if not m:
            break

        inner = normalize_name(m.group(1))
        value = lookup.get(inner, '')

        s, e = m.start(), m.end()

        offsets = []
        pos = 0
        for t in texts:
            offsets.append(pos)
            pos += len(t)

        first_i = None
        for i in range(len(texts)):
            if offsets[i] <= s < offsets[i] + len(texts[i]):
                first_i = i
                break
        last_i = None
        for j in range(len(texts)):
            if offsets[j] < e <= offsets[j] + len(texts[j]):
                last_i = j
        if first_i is None:
            break
        if last_i is None:
            last_i = len(texts) - 1

        prefix = texts[first_i][:s - offsets[first_i]]
        suffix = texts[last_i][e - offsets[last_i]:]

        if first_i == last_i:
            runs[first_i].text = prefix + value + suffix
        else:
            runs[first_i].text = prefix + value
            for k in range(first_i + 1, last_i):
                runs[k].text = ''
            runs[last_i].text = suffix


# ---------- Обхід усіх абзаців документа (тіло, таблиці, колонтитули) ----------

def collect_paragraphs(container):
    paras = []
    try:
        paras.extend(container.paragraphs)
    except Exception:
        pass
    try:
        for table in container.tables:
            for row in table.rows:
                for cell in row.cells:
                    paras.extend(collect_paragraphs(cell))
    except Exception:
        pass
    return paras


def fill_document(docx_bytes, data_map):
    pattern, lookup = build_matcher(data_map)
    doc = Document(io.BytesIO(docx_bytes))

    if pattern is not None:
        for p in collect_paragraphs(doc):
            replace_in_paragraph(p, pattern, lookup)

        for section in doc.sections:
            for hf in (section.header, section.first_page_header,
                       section.even_page_header, section.footer,
                       section.first_page_footer, section.even_page_footer):
                try:
                    for p in collect_paragraphs(hf):
                        replace_in_paragraph(p, pattern, lookup)
                except Exception:
                    pass

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ---------- Маршрути ----------

@app.route('/', methods=['GET'])
def root():
    return 'DOCX filler service is running.', 200


@app.route('/health', methods=['GET'])
def health():
    return 'ok', 200


@app.route('/generate', methods=['POST'])
def generate():
    try:
        job = request.get_json(force=True, silent=False)
    except Exception as e:
        return ('Bad JSON: ' + str(e), 400)

    if not job or 'people' not in job:
        return ('Missing "people" in request', 400)

    mem = io.BytesIO()
    with zipfile.ZipFile(mem, 'w', zipfile.ZIP_DEFLATED) as zf:
        for person in job.get('people', []):
            folder = person.get('folder', '') or ''
            data_map = person.get('data', {}) or {}
            for f in person.get('files', []):
                out_name = f.get('outName', 'file.docx')
                b64 = f.get('templateB64', '')
                if not b64:
                    continue
                try:
                    template_bytes = base64.b64decode(b64)
                    filled = fill_document(template_bytes, data_map)
                except Exception as e:
                    # кладемо .txt із помилкою замість зламаного docx,
                    # щоб було видно, який файл не вдався
                    err = ('Помилка обробки цього файлу: ' + str(e)).encode('utf-8')
                    path = (folder + '/' if folder else '') + out_name + '.ERROR.txt'
                    zf.writestr(path, err)
                    continue
                path = (folder + '/' if folder else '') + out_name
                zf.writestr(path, filled)

    data = mem.getvalue()
    return Response(
        data,
        mimetype='application/zip',
        headers={'Content-Disposition': 'attachment; filename="result.zip"'}
    )


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
