#!/usr/bin/env python3
"""DOCX → bullets HTML + entrega DOCX. Solo biblioteca estándar de Python."""
import argparse
from copy import deepcopy
import hashlib
from html import escape
from html.parser import HTMLParser
import json
from pathlib import Path
import re
import sys
import unicodedata
from xml.etree import ElementTree as ET
from zipfile import ZipFile

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}
ET.register_namespace('w', W)
ASSETS = Path(__file__).resolve().parents[1] / 'assets'


def q(name):
    return '{' + W + '}' + name


def norm(text):
    return ''.join(c for c in unicodedata.normalize('NFD', text.casefold())
                   if unicodedata.category(c) != 'Mn')


def paragraph_text(p):
    parts = []
    for node in p.iter():
        if node.tag == q('t'):
            parts.append(node.text or '')
        elif node.tag in (q('br'), q('cr')):
            parts.append('\n')
        elif node.tag == q('tab'):
            parts.append('\t')
    return ''.join(parts).strip()


def capitalize_first(text):
    text = text.strip()
    for i, char in enumerate(text):
        if char.isalpha():
            return text[:i] + char.upper() + text[i + 1:]
    return text


class BulletHTML(HTMLParser):
    """Lee listas literales incluso si sus tags están partidos entre runs."""
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.depth = 0
        self.current = None
        self.items = []
        self.forbidden = False

    def handle_starttag(self, tag, attrs):
        if tag == 'ul':
            self.depth += 1
            if self.depth > 1:
                raise ValueError('Lista HTML anidada: requiere decidir cómo representar la jerarquía.')
        elif tag == 'li':
            if self.depth != 1 or self.current is not None:
                raise ValueError('HTML de lista incompleto o ambiguo.')
            self.current = []
        elif tag == 'br' and self.current is not None:
            self.current.append(' ')
        elif tag in ('script', 'style', 'img', 'ol', 'table'):
            self.forbidden = True

    def handle_endtag(self, tag):
        if tag == 'li':
            if self.current is None:
                raise ValueError('Cierre </li> sin apertura.')
            text = ''.join(self.current).strip()
            if not text:
                raise ValueError('Bullet HTML vacío.')
            self.items.append(re.sub(r'[\r\n\t]+', ' ', text))
            self.current = None
        elif tag == 'ul':
            if self.current is not None or self.depth != 1:
                raise ValueError('Falta cerrar un elemento de la lista HTML.')
            self.depth -= 1

    def handle_data(self, data):
        if self.current is not None:
            self.current.append(data)
        elif data.strip():
            raise ValueError('Texto fuera de <li> dentro del bloque HTML; revisar origen.')

    def result(self):
        self.close()
        if self.depth or self.current is not None or not self.items or self.forbidden:
            raise ValueError('Lista HTML incompleta o con contenido no compatible.')
        return self.items


class Numbering:
    def __init__(self, archive):
        self.styles = {}
        self.nums = {}
        self.abstracts = {}
        if 'word/styles.xml' in archive.namelist():
            root = ET.fromstring(archive.read('word/styles.xml'))
            self.styles = {s.get(q('styleId')): s for s in root.findall('w:style', NS)}
        if 'word/numbering.xml' in archive.namelist():
            root = ET.fromstring(archive.read('word/numbering.xml'))
            self.nums = {n.get(q('numId')): n for n in root.findall('w:num', NS)}
            self.abstracts = {n.get(q('abstractNumId')): n
                              for n in root.findall('w:abstractNum', NS)}

    def bullet_level(self, p):
        values = {}
        pr = p.find('w:pPr', NS)
        seen = set()
        while pr is not None:
            for key in ('numId', 'ilvl'):
                node = pr.find('w:numPr/w:' + key, NS)
                if node is not None and key not in values:
                    values[key] = node.get(q('val'))
            style = pr.find('w:pStyle', NS)
            name = style.get(q('val')) if style is not None else None
            if not name or name in seen:
                break
            seen.add(name)
            s = self.styles.get(name)
            if s is None:
                break
            pr = deepcopy(s.find('w:pPr', NS))
            parent = s.find('w:basedOn', NS)
            if parent is not None:
                if pr is None:
                    pr = ET.Element(q('pPr'))
                ET.SubElement(pr, q('pStyle'), {q('val'): parent.get(q('val'))})
        num_id = values.get('numId')
        if not num_id or num_id == '0':
            return None
        level = values.get('ilvl', '0')
        n = self.nums.get(num_id)
        if n is None:
            raise ValueError(f'No existe la definición de lista numId={num_id}.')
        lvl = n.find(f'w:lvlOverride[@w:ilvl="{level}"]/w:lvl', NS)
        if lvl is None:
            aid = n.find('w:abstractNumId', NS)
            a = self.abstracts.get(aid.get(q('val'))) if aid is not None else None
            lvl = a.find(f'w:lvl[@w:ilvl="{level}"]', NS) if a is not None else None
        fmt = lvl.find('w:numFmt', NS) if lvl is not None else None
        if fmt is None:
            raise ValueError(f'No se puede resolver el formato de lista {num_id}/{level}.')
        return int(level) if fmt.get(q('val')) == 'bullet' else None


def extract(path):
    with ZipFile(path) as archive:
        root = ET.fromstring(archive.read('word/document.xml'))
        if any(root.find('.//w:' + tag, NS) is not None for tag in ('ins', 'del', 'altChunk')):
            raise ValueError('Documento con revisiones o contenido incrustado; resolver antes de extraer.')
        numbering = Numbering(archive)
        paragraphs = root.findall('.//w:body//w:p', NS)
        rows = [(i, paragraph_text(p), numbering.bullet_level(p))
                for i, p in enumerate(paragraphs)]
    rows = [row for row in rows if row[1]]
    if not rows:
        raise ValueError('Documento vacío.')
    title = rows[0][1]
    if rows[0][2] is not None or '<ul' in title.lower():
        raise ValueError('No se identificó un título antes de la lista.')
    sections = []
    prev = ''
    i = 1
    while i < len(rows):
        index, text, level = rows[i]
        start = index
        raw = None
        kind = None
        if re.search(r'<ul\b', text, re.I):
            chunks = [text]
            while not re.search(r'</ul\s*>', chunks[-1], re.I):
                i += 1
                if i >= len(rows):
                    raise ValueError('Falta </ul> en el DOCX.')
                chunks.append(rows[i][1])
            parser = BulletHTML()
            parser.feed('\n'.join(chunks))
            raw = parser.result()
            kind = 'html'
        elif level is not None:
            raw = []
            while i < len(rows) and rows[i][2] is not None:
                if rows[i][2] != 0:
                    raise ValueError('Lista nativa con niveles: requiere resolver la jerarquía.')
                raw.append(re.sub(r'[\r\n\t]+', ' ', rows[i][1]))
                i += 1
            i -= 1
            kind = 'native'
        elif re.search(r'</?(?:li|ul)\b', text, re.I):
            raise ValueError('Fragmento HTML de lista sin un <ul> completo.')
        if raw is not None:
            if not prev:
                raise ValueError('No hay subtítulo inmediatamente antes de la lista.')
            sections.append({'subtitle_original': prev, 'raw': raw,
                             'bullets': [capitalize_first(x) for x in raw],
                             'kind': kind, 'paragraph_start': start})
            prev = ''
        else:
            prev = text
        i += 1
    if not sections:
        raise ValueError('No se encontraron bullets nativos ni <ul><li>.')
    return {'source': path.name, 'sha256': hashlib.sha256(path.read_bytes()).hexdigest(),
            'title': title, 'sections': sections}


def token_match(text, term):
    return re.search(r'(?<![a-z0-9])' + re.escape(norm(term)) + r'(?![a-z0-9])', norm(text))


def university_for(filename, config, explicit=None):
    matches = [key for key, profile in config['universities'].items()
               if any(token_match(filename, alias) for alias in profile['filename_aliases'])]
    if len(matches) > 1 or (explicit and matches and explicit not in matches):
        raise ValueError(f'Universidad ambigua o contradictoria: {matches}.')
    if explicit:
        if explicit not in config['universities']:
            raise ValueError(f'Falta el perfil visual de {explicit}. Añadirlo a perfiles.json.')
        return explicit
    if len(matches) != 1:
        raise ValueError('Universidad desconocida: añadir perfil y alias; no se heredan colores.')
    return matches[0]


def degree_for(filename, title, profile, config, explicit=None):
    matches = {key for key, aliases in config['degrees'].items()
               if any(token_match(filename, alias) for alias in aliases)}
    matches.update(value for alias, value in profile.get('filename_degrees', {}).items()
                   if token_match(filename, alias))
    # El título solo confirma el código del nombre o sirve como respaldo documentado.
    title_matches = {key for key, aliases in config['degrees'].items()
                     if any(norm(title).startswith(norm(alias) + ' ') or norm(title) == norm(alias)
                            for alias in aliases)}
    if len(matches) > 1 or len(title_matches) > 1 or (matches and title_matches and matches != title_matches):
        raise ValueError(f'Grado ambiguo: nombre={sorted(matches)}, título={sorted(title_matches)}.')
    detected = matches or title_matches
    if explicit:
        if detected and explicit not in detected:
            raise ValueError(f'Grado explícito {explicit} contradice {sorted(detected)}.')
        return explicit, 'explicit'
    if len(detected) != 1:
        raise ValueError('Grado desconocido: usar --degree con su etiqueta o registrar alias.')
    return next(iter(detected)), 'filename' if matches else 'title'


def safe_label(value):
    if not re.fullmatch(r'[\w-]+', value, re.UNICODE):
        raise ValueError(f'Etiqueta de archivo no válida: {value!r}. Usar letras, números o guiones.')
    return value


def set_css(tag, prop, value):
    style = re.search(r'\bstyle="([^"]*)"', tag)
    if not style:
        raise ValueError('La plantilla requiere estilos inline en li y span.')
    css, count = re.subn(r'(?<![\w-])(' + re.escape(prop) + r'\s*:\s*)[^;]+',
                         lambda m: m[1] + value, style[1])
    if count != 1:
        raise ValueError(f'Propiedad {prop} ausente o repetida en plantilla.')
    return tag[:style.start(1)] + css + tag[style.end(1):]


def render_html(base, bullets, profile):
    for key in ('text_color', 'bullet_color'):
        if not re.fullmatch(r'#[0-9a-fA-F]{6}', profile[key]):
            raise ValueError(f'Color inválido: {key}.')
    if not re.fullmatch(r'\d+(?:\.\d+)?(?:px|pt)', profile['text_size']):
        raise ValueError('text_size debe incluir px o pt.')
    lists = list(re.finditer(r'(<ul\b[^>]*>)([\s\S]*?)(</ul>)', base, re.I))
    if len(lists) != 1:
        raise ValueError('BASE_HTML debe contener exactamente un ul.')
    ul = lists[0]
    items = re.findall(r'<li\b[\s\S]*?</li>', ul[2], re.I)
    if not items:
        raise ValueError('BASE_HTML no contiene un li de referencia.')
    fragments = []
    for i, text in enumerate(bullets):
        item = items[i] if i < len(items) else items[0]
        item = re.sub(r'<li\b[^>]*>', lambda m: set_css(m[0], 'color', profile['bullet_color']), item, count=1)
        item, count = re.subn(r'(<span\b[^>]*>)[\s\S]*?(</span>)',
                             lambda m: set_css(set_css(m[1], 'color', profile['text_color']),
                                               'font-size', profile['text_size']) + escape(text, quote=False) + m[2],
                             item, count=1, flags=re.I)
        if count != 1:
            raise ValueError('Cada li de BASE_HTML debe tener un span.')
        fragments.append('        ' + item)
    # Se conservan atributos, clases, fuentes, line-height, márgenes y contenedor.
    return (base[:ul.start(2)] + '\n' + '\n'.join(fragments) + '\n      ' + base[ul.end(2):]).strip() + '\n'


def word_paragraph(text, kind):
    p = ET.Element(q('p'))
    pr = ET.SubElement(p, q('pPr'))
    ET.SubElement(pr, q('pStyle'), {q('val'): 'Normal'})
    if kind in ('title', 'subtitle'):
        ET.SubElement(pr, q('keepNext'))
    if kind == 'title':
        ET.SubElement(pr, q('spacing'), {q('beforeAutospacing'): '1', q('afterAutospacing'): '1'})
        ET.SubElement(pr, q('outlineLvl'), {q('val'): '2'})
    run = ET.SubElement(p, q('r'))
    rp = ET.SubElement(run, q('rPr'))
    ET.SubElement(rp, q('rFonts'), {q('ascii'): 'Aptos', q('hAnsi'): 'Aptos', q('eastAsia'): 'Aptos'})
    if kind in ('title', 'subtitle'):
        ET.SubElement(rp, q('b'))
        ET.SubElement(rp, q('bCs'))
    else:
        for flag in ('b', 'bCs', 'i', 'iCs'):
            ET.SubElement(rp, q(flag), {q('val'): '0'})
    ET.SubElement(rp, q('color'), {q('val'): '000000'})
    for size_tag in ('sz', 'szCs'):
        ET.SubElement(rp, q(size_tag), {q('val'): '18' if kind == 'code' else '24'})
    if kind in ('title', 'subtitle'):
        ET.SubElement(rp, q('highlight'), {q('val'): 'yellow' if kind == 'title' else 'green'})
    else:
        ET.SubElement(rp, q('highlight'), {q('val'): 'none'})
        ET.SubElement(rp, q('u'), {q('val'): 'none'})
    ET.SubElement(run, q('t'), {'{http://www.w3.org/XML/1998/namespace}space': 'preserve'}).text = text
    return ET.tostring(p, encoding='unicode')


def make_docx(template, destination, records):
    paragraphs = []
    for record in records:
        paragraphs.append(word_paragraph('', 'blank'))
        paragraphs.append(word_paragraph(record['title'], 'title'))
        for section in record['sections']:
            paragraphs.append(word_paragraph(section['subtitle'], 'subtitle'))
            paragraphs.append(word_paragraph('', 'blank'))
            paragraphs.extend(word_paragraph(line, 'code') for line in section['html'].splitlines())
    with ZipFile(template) as source, ZipFile(destination, 'w') as dest:
        for entry in source.infolist():
            data = source.read(entry.filename)
            if entry.filename == 'word/document.xml':
                xml = data.decode('utf-8')
                sect = re.search(r'<w:sectPr\b[\s\S]*?</w:sectPr>', xml)
                if not sect:
                    raise ValueError('Plantilla sin configuración de página w:sectPr.')
                body = ''.join(paragraphs) + sect[0]
                xml, count = re.subn(r'(<w:body\b[^>]*>)[\s\S]*?(</w:body>)',
                                     lambda m: m[1] + body + m[2], xml)
                if count != 1:
                    raise ValueError('Plantilla Word no compatible.')
                ET.fromstring(xml)
                data = xml.encode('utf-8')
            dest.writestr(entry, data)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--input', required=True, type=Path, help='Carpeta de DOCX originales (sin recursión)')
    parser.add_argument('--output', required=True, type=Path, help='Carpeta nueva para esta entrega')
    parser.add_argument('--config', type=Path, default=ASSETS / 'perfiles.json')
    parser.add_argument('--base', type=Path, default=ASSETS / 'base.html')
    parser.add_argument('--script', type=Path, help='Leer BASE_HTML de script.js, sin ejecutarlo')
    parser.add_argument('--template', type=Path, default=ASSETS / 'plantilla.docx')
    parser.add_argument('--university', help='Id de perfil para nombres sin universidad reconocible')
    parser.add_argument('--degree', help='Etiqueta explícita para un grado nuevo, p.ej. Certificaciones')
    parser.add_argument('--literal-subtitles', action='store_true')
    args = parser.parse_args()
    if args.output.exists():
        raise ValueError('La carpeta de salida ya existe. Elige una nueva para conservar entregas anteriores.')
    files = sorted((p for p in args.input.iterdir() if p.suffix.lower() == '.docx'
                    and not p.name.startswith('~$')), key=lambda p: norm(p.name))
    if not files:
        raise ValueError('No hay documentos DOCX en la carpeta de entrada.')
    config = json.loads(args.config.read_text())
    if args.script:
        match = re.search(r'const\s+BASE_HTML\s*=\s*String\.raw`([^`]*)`', args.script.read_text())
        if not match:
            raise ValueError('No se encontró const BASE_HTML = String.raw`...` en script.js.')
        base = match[1].strip()
    else:
        base = args.base.read_text().strip()
    groups = {}
    errors = []
    for path in files:
        try:
            record = extract(path)
            university = university_for(path.stem, config, args.university)
            profile = config['universities'][university]
            degree, evidence = degree_for(path.stem, record['title'], profile, config, args.degree)
            safe_label(university)
            safe_label(degree)
            record.update(university=university, degree=degree, degree_evidence=evidence)
            for section in record['sections']:
                original = section['subtitle_original']
                section['subtitle'] = original if args.literal_subtitles else config.get('subtitle_replacements', {}).get(original, original)
                section['html'] = render_html(base, section['bullets'], profile)
            groups.setdefault((university, degree), []).append(record)
        except (ValueError, KeyError, ET.ParseError) as error:
            errors.append(f'{path.name}: {error}')
    if errors:
        raise ValueError('No se generó una entrega parcial. Resolver:\n' + '\n'.join(errors))
    # Validar plantilla antes de crear la carpeta de salida.
    with ZipFile(args.template) as archive:
        template_xml = archive.read('word/document.xml').decode()
        if not re.search(r'<w:sectPr\b[\s\S]*?</w:sectPr>', template_xml):
            raise ValueError('Plantilla sin configuración de página compatible.')
    args.output.mkdir(parents=True)
    preview = ['<!doctype html><html lang="es"><meta charset="utf-8"><title>Revisión de bullets</title>',
               '<style>body{font:16px Arial;background:#eee;margin:32px auto;max-width:960px;padding:0 20px}'
               'article{background:white;padding:24px;margin:24px 0}pre{white-space:pre-wrap;overflow-wrap:anywhere;'
               'font-size:12px;background:#f5f5f5;padding:16px}h2{font-size:22px}h3{font-size:18px}</style>',
               '<h1>Revisión de bullets</h1><p>Vista previa y fragmentos HTML de cada sección.</p>']
    report = {'input': str(args.input.resolve()), 'files': len(files), 'deliveries': []}
    for (university, degree), records in sorted(groups.items()):
        stem = f'{university}_{len(records)}_{degree}'
        folder = args.output / stem
        folder.mkdir()
        for n, record in enumerate(records, 1):
            for j, section in enumerate(record['sections'], 1):
                section['html_file'] = f'{stem}/{n:02d}_{j:02d}.html'
                section['raw_file'] = f'{stem}/{n:02d}_{j:02d}.txt'
                (args.output / section['html_file']).write_text(section['html'])
                (args.output / section['raw_file']).write_text('\n'.join(section['bullets']) + '\n')
                preview.extend(['<article>', '<h2>' + escape(record['title']) + '</h2>',
                                '<h3>' + escape(section['subtitle']) + '</h3>', section['html'],
                                '<details><summary>Ver HTML formateado</summary><pre>' + escape(section['html']) + '</pre></details></article>'])
        docx = args.output / (stem + '.docx')
        make_docx(args.template, docx, records)
        report['deliveries'].append({'docx': docx.name, 'documents': len(records),
                                    'sections': sum(len(r['sections']) for r in records),
                                    'bullets': sum(len(s['bullets']) for r in records for s in r['sections']),
                                    'records': records})
    preview.append('</html>')
    (args.output / 'revision.html').write_text('\n'.join(preview))
    (args.output / 'manifest.json').write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n')
    for delivery in report['deliveries']:
        print(f"{delivery['docx']}: {delivery['documents']} documentos, {delivery['sections']} secciones, {delivery['bullets']} bullets")
    print('Vista previa:', args.output / 'revision.html')


if __name__ == '__main__':
    try:
        main()
    except (ValueError, OSError, KeyError, ET.ParseError) as error:
        sys.exit(str(error))
