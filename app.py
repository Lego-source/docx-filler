import base64
import io
import re
import zipfile
import copy
from flask import Flask, request, Response

from docx import Document
from docx.text.paragraph import Paragraph

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024


# ---------- Нормалізація ----------

def normalize_name(s):
    s = (s or '')
    s = s.replace('_', ' ').replace('\u00A0', ' ')
    s = re.sub(r'\s+', ' ', s).strip()
    return s.lower()


def normalize_apostrophes_(s):
    return (s or '').replace('\u2019', "'").replace('\u02bc', "'")


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


# ---------- Вставка нового абзацу після якоря (спільна допоміжна) ----------

def insert_paragraph_after(anchor_paragraph, text):
    new_p = copy.deepcopy(anchor_paragraph._p)
    anchor_paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, anchor_paragraph._parent)
    runs = new_para.runs
    if runs:
        runs[0].text = text
        for r in runs[1:]:
            r.text = ''
    else:
        new_para.add_run(text)
    return new_para


# ---------- ТВО-вставка (без змін) ----------

TVO_A_MARK   = 'прошу покласти тимчасове виконання'
TVO_B_START  = 'не заперечую'
TVO_B_END    = 'тво_фам'
ANCHOR_ADDR  = 'відпустку буду проводити за адресою'


def paragraph_plain_text(p):
    return ''.join(r.text for r in p.runs)


def get_tvo_parts(donor_doc):
    paras = donor_doc.paragraphs
    a_idx = None
    for i, p in enumerate(paras):
        if TVO_A_MARK in paragraph_plain_text(p).lower():
            a_idx = i
            break
    if a_idx is None:
        return None, None
    b_start = None
    for i in range(a_idx + 1, len(paras)):
        if TVO_B_START in paragraph_plain_text(paras[i]).lower():
            b_start = i
            break
    if b_start is None:
        return None, None
    b_end = None
    for i in range(b_start, len(paras)):
        if TVO_B_END in paragraph_plain_text(paras[i]).lower().replace(' ', ''):
            b_end = i
            break
    if b_end is None:
        return None, None
    part_a = [copy.deepcopy(paras[a_idx]._p)]
    part_b = [copy.deepcopy(paras[i]._p) for i in range(b_start, b_end + 1)]
    return part_a, part_b


def find_addr_anchor(doc):
    for p in doc.paragraphs:
        if ANCHOR_ADDR in paragraph_plain_text(p).lower():
            return p
    return None


def is_applicant_signature(p):
    t = paragraph_plain_text(p).lower()
    return ('\u00abзван\u00bb' in t) and ('\u00abім\u00bb' in t or '\u00abiм\u00bb' in t) and ('\u00abфам\u00bb' in t)


def find_applicant_signature(doc):
    found = None
    for p in doc.paragraphs:
        if is_applicant_signature(p):
            found = p
    return found


def insert_tvo_block(target_doc, donor_doc):
    part_a, part_b = get_tvo_parts(donor_doc)
    if not part_a or not part_b:
        return False, 'donor'
    addr = find_addr_anchor(target_doc)
    if addr is None:
        return False, 'anchor_addr'
    sign = find_applicant_signature(target_doc)
    if sign is None:
        return False, 'anchor_sign'
    ref = sign._p
    for para_xml in part_b:
        ref.addnext(para_xml)
        ref = para_xml
    ref = addr._p
    for para_xml in part_a:
        ref.addnext(para_xml)
        ref = para_xml
    return True, None


# ---------- НОВЕ: пункт 8 контракту для військовослужбовців 45+ ----------

AGE_CLAUSE_ANCHOR = 'перебування у стані алкогольного'
AGE_CLAUSE_TEXT = (
    "Після закінчення особливого періоду, у разі досягнення військовослужбовцем "
    "граничного віку перебування на військовій службі, контракт припиняється "
    "(розривається), а військовослужбовець підлягає звільненню з військової служби за віком."
)


def insert_age_clause(doc):
    """Шукає абзац з ознакою пункту 8 («...стані алкогольного...») і дописує
       після нього додаткове речення. Повертає True, якщо знайдено й вставлено."""
    for p in doc.paragraphs:
        text = normalize_apostrophes_(''.join(r.text for r in p.runs)).lower()
        if AGE_CLAUSE_ANCHOR in text:
            insert_paragraph_after(p, AGE_CLAUSE_TEXT)
            return True
    return False


# ---------- Заповнення одного документа ----------

def fill_document(docx_bytes, data_map, tvo_donor_bytes=None, add_tvo=False, add_age_clause=False):
    doc = Document(io.BytesIO(docx_bytes))

    tvo_note = None
    if add_tvo:
        if tvo_donor_bytes is None:
            tvo_note = 'ТВО: не передано шаблон-донор (№2).'
        else:
            donor = Document(io.BytesIO(tvo_donor_bytes))
            ok, why = insert_tvo_block(doc, donor)
            if not ok:
                notes = {
                    'donor': 'ТВО: у доноровому шаблоні №2 не знайдено потрібні абзаци.',
                    'anchor_addr': 'ТВО: не знайдено абзац «Відпустку буду проводити за адресою…».',
                    'anchor_sign': 'ТВО: не знайдено підпис заявника («зван» «ім» «фам») для вставки блоку «Не заперечую».'
                }
                tvo_note = notes.get(why, 'ТВО: не вдалося вставити блок.')

    age_note = None
    if add_age_clause:
        found = insert_age_clause(doc)
        if not found:
            # М'яке попередження: людині 45+, але в ЦЬОМУ конкретному файлі
            # не знайдено пункт 8 — можливо, у нього просто немає такого абзацу
            # (наприклад, це не сам текст контракту, а супровідний документ).
            age_note = ('Вік 45+: пункт 8 не знайдено в цьому файлі '
                        '(очікувався абзац «...перебування у стані алкогольного...»).')

    apply_placeholders(doc, data_map)

    out = io.BytesIO()
    doc.save(out)

    note = None
    if tvo_note and age_note:
        note = tvo_note + '\n' + age_note
    else:
        note = tvo_note or age_note

    return out.getvalue(), note


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
            add_age_clause = bool(person.get('addAgeClause', False))
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
                        add_tvo=file_add_tvo,
                        add_age_clause=add_age_clause
                    )
                except Exception as e:
                    err = ('Помилка обробки: ' + str(e)).encode('utf-8')
                    path = (folder + '/' if folder else '') + out_name + '.ERROR.txt'
                    zf.writestr(path, err)
                    continue
                path = (folder + '/' if folder else '') + out_name
                zf.writestr(path, filled)
                if note:
                    zf.writestr(path + '.УВАГА.txt', note.encode('utf-8'))

    return Response(
        mem.getvalue(),
        mimetype='application/zip',
        headers={'Content-Disposition': 'attachment; filename="result.zip"'}
    )


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
