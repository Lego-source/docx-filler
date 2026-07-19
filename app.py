import base64
import io
import re
import zipfile
import copy
from flask import Flask, request, Response

from docx import Document

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024


# ---------- Нормалізація ----------

def normalize_name(s):
    s = (s or '')
    s = s.replace('_', ' ').replace('\u00A0', ' ')
    s = re.sub(r'\s+', ' ', s).strip()
    return s.lower()


def build_matcher(data_map):
    lookup = {}
    names = []
    for key, val in data_map.items():
        norm = normalize_name(key)
        if norm:
            lookup[norm] = '' if val is None else str(val)
            names.append(key)
    if not names:
        return None, {}
    names.sort(key=lambda n: len(n), reverse=True)
    alts = []
    for n in names:
        piece = re.escape(n.strip())
        piece = re.sub(r'\\?\s+', r'[\\s_]+', piece)
        alts.append(piece)
    pattern = re.compile(
        '\u00AB[\\s_]*(' + '|'.join(alts) + ')[\\s_]*\u00BB',
        re.IGNORECASE | re.UNICODE
    )
    return pattern, lookup


# ---------- Заміна в абзаці ----------

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


def apply_placeholders(doc, data_map):
    pattern, lookup = build_matcher(data_map)
    if pattern is None:
        return
    for p in collect_paragraphs(doc):
        replace_in_paragraph(p, pattern, lookup)
    for section in doc.sections:
        for hf in (section.header, section.first_page_header, section.even_page_header,
                   section.footer, section.first_page_footer, section.even_page_footer):
            try:
                for p in collect_paragraphs(hf):
                    replace_in_paragraph(p, pattern, lookup)
            except Exception:
                pass


# ---------- ТВО-вставка ----------

# У шаблоні-донорі (№2) ТВО-блок починається з цієї фрази і закінчується
# абзацом підпису ТВО, що містить «ТВО_ім» «ТВО_фам».
TVO_START_MARK = 'прошу покласти тимчасове виконання'
TVO_END_MARK   = 'тво_фам'

# У ЦІЛЬОВОМУ шаблоні вставляємо блок ПІСЛЯ абзацу з адресою відпустки,
# тобто ПЕРЕД підписом заявника.
ANCHOR_MARK = 'відпустку буду проводити за адресою'


def paragraph_plain_text(p):
    return ''.join(r.text for r in p.runs)


def find_tvo_block_paragraphs(donor_doc):
    paras = donor_doc.paragraphs
    start = None
    end = None
    for i, p in enumerate(paras):
        low = paragraph_plain_text(p).lower()
        if start is None and TVO_START_MARK in low:
            start = i
        if start is not None and TVO_END_MARK in low.replace(' ', ''):
            end = i
            break
    if start is None or end is None or end < start:
        return None
    return [copy.deepcopy(paras[i]._p) for i in range(start, end + 1)]


def find_anchor_paragraph(doc):
    """Абзац з адресою відпустки — вставляємо ТВО одразу після нього."""
    for p in doc.paragraphs:
        if ANCHOR_MARK in paragraph_plain_text(p).lower():
            return p
    return None


def insert_tvo_block(target_doc, donor_doc):
    block = find_tvo_block_paragraphs(donor_doc)
    if not block:
        return False, 'donor'
    anchor = find_anchor_paragraph(target_doc)
    if anchor is None:
        return False, 'anchor'
    ref = anchor._p
    for para_xml in block:
        ref.addnext(para_xml)
        ref = para_xml
    return True, None


def fill_document(docx_bytes, data_map, tvo_donor_bytes=None, add_tvo=False):
    doc = Document(io.BytesIO(docx_bytes))
    tvo_note = None
    if add_tvo:
        if tvo_donor_bytes is None:
            tvo_note = 'ТВО: не передано шаблон-донор (№2).'
        else:
            donor = Document(io.BytesIO(tvo_donor_bytes))
            ok, why = insert_tvo_block(doc, donor)
            if not ok:
                if why == 'donor':
                    tvo_note = ('ТВО: у доноровому шаблоні №2 не знайдено блок '
                                '(«Прошу покласти тимчасове виконання…» … «ТВО_ім» «ТВО_фам»).')
                else:
                    tvo_note = ('ТВО: у цьому шаблоні не знайдено абзац-якір '
                                '«Відпустку буду проводити за адресою…» для вставки блоку.')
    apply_placeholders(doc, data_map)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue(), tvo_note


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
        return ('Missing "people"', 400)

    mem = io.BytesIO()
    with zipfile.ZipFile(mem, 'w', zipfile.ZIP_DEFLATED) as zf:
        for person in job.get('people', []):
            folder = person.get('folder', '') or ''
            data_map = person.get('data', {}) or {}
            add_tvo = bool(person.get('addTvo', False))
            tvo_donor_b64 = person.get('tvoDonorB64')
            tvo_donor_bytes = base64.b64decode(tvo_donor_b64) if tvo_donor_b64 else None

            for f in person.get('files', []):
                out_name = f.get('outName', 'file.docx')
                b64 = f.get('templateB64', '')
                file_add_tvo = add_tvo and bool(f.get('allowTvo', True))
                if not b64:
                    continue
                try:
                    template_bytes = base64.b64decode(b64)
                    filled, note = fill_document(
                        template_bytes, data_map,
                        tvo_donor_bytes=tvo_donor_bytes,
                        add_tvo=file_add_tvo
                    )
                except Exception as e:
                    err = ('Помилка обробки: ' + str(e)).encode('utf-8')
                    path = (folder + '/' if folder else '') + out_name + '.ERROR.txt'
                    zf.writestr(path, err)
                    continue
                path = (folder + '/' if folder else '') + out_name
                zf.writestr(path, filled)
                if note:
                    zf.writestr(path + '.ТВО-увага.txt', note.encode('utf-8'))

    return Response(
        mem.getvalue(),
        mimetype='application/zip',
        headers={'Content-Disposition': 'attachment; filename="result.zip"'}
    )


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
