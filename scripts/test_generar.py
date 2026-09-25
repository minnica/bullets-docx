import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import generar as g


class PipelineTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.config = json.loads((g.ASSETS / 'perfiles.json').read_text())
        self.profile = self.config['universities']['IBERO']
        self.base = (g.ASSETS / 'base.html').read_text()

    def source(self, paragraphs, numbering='', styles=''):
        file = self.root / 'Ibero_PUJ_D_prueba.docx'
        with ZipFile(file, 'w') as z:
            z.writestr('word/document.xml', f'<w:document xmlns:w="{g.W}"><w:body>{paragraphs}</w:body></w:document>')
            if numbering:
                z.writestr('word/numbering.xml', numbering)
            if styles:
                z.writestr('word/styles.xml', styles)
        return file

    def test_html_runs_punctuation_case_and_entities(self):
        paragraphs = g.word_paragraph('Diplomado de Prueba', 'title') + g.word_paragraph('Beneficios', 'subtitle')
        paragraphs += g.word_paragraph('<ul><li>¿aprender IA &amp; UX?</li><li>usar &lt;datos&gt;</li></ul>', 'code')
        record = g.extract(self.source(paragraphs))
        self.assertEqual(record['sections'][0]['bullets'], ['¿Aprender IA & UX?', 'Usar <datos>'])
        html = g.render_html(self.base, record['sections'][0]['bullets'], self.profile)
        root = ET.fromstring(html)
        self.assertEqual([x.text for x in root.findall('.//span')], ['¿Aprender IA & UX?', 'Usar <datos>'])

    def test_native_lists_inherited_styles_and_numbered_exclusion(self):
        numbering = f'''<w:numbering xmlns:w="{g.W}">
        <w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl></w:abstractNum>
        <w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>
        <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
        <w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num></w:numbering>'''
        styles = f'''<w:styles xmlns:w="{g.W}"><w:style w:styleId="Parent"><w:pPr><w:numPr>
        <w:numId w:val="1"/></w:numPr></w:pPr></w:style><w:style w:styleId="Child"><w:basedOn w:val="Parent"/></w:style></w:styles>'''
        paragraphs = g.word_paragraph('Diplomado de Prueba', 'title') + g.word_paragraph('Competencias', 'subtitle')
        paragraphs += '<w:p><w:pPr><w:pStyle w:val="Child"/></w:pPr><w:r><w:t>aprender IA.</w:t><w:br/><w:t>Sin perder.</w:t></w:r></w:p>'
        paragraphs += '<w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>crear sin punto</w:t></w:r></w:p>'
        paragraphs += '<w:p><w:pPr><w:numPr><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>Excluir numeración</w:t></w:r></w:p>'
        record = g.extract(self.source(paragraphs, numbering, styles))
        self.assertEqual(len(record['sections']), 1)
        self.assertEqual(record['sections'][0]['bullets'], ['Aprender IA. Sin perder.', 'Crear sin punto'])

    def test_template_structure_changes_only_allowed_properties(self):
        profile = dict(self.profile, text_color='#123456', bullet_color='#ABCDEF', text_size='18px')
        output = ET.fromstring(g.render_html(self.base, ['Prueba.'] * 9, profile))
        original = ET.fromstring(self.base)
        self.assertEqual(output.attrib, original.attrib)
        self.assertEqual(len(output.findall('.//li')), 9)
        for li in output.findall('.//li'):
            self.assertEqual(li.attrib['class'], original.find('.//li').attrib['class'])
            self.assertEqual(li.attrib['style'], original.find('.//li').attrib['style'].replace('#E00034', '#ABCDEF'))
            self.assertEqual(li.find('span').attrib['style'], original.find('.//span').attrib['style'].replace('#2E2E2E', '#123456').replace('16px', '18px'))

    def test_unknown_university_and_degree_require_information(self):
        with self.assertRaises(ValueError):
            g.university_for('OTRA_D_documento', self.config)
        with self.assertRaises(ValueError):
            g.degree_for('Ibero_X_programa', 'Certificación de Prueba', self.profile, self.config)
        self.assertEqual(g.degree_for('Ibero_X_programa', 'Certificación de Prueba', self.profile, self.config, 'Cert')[0], 'Cert')
        self.assertEqual(g.degree_for('Ibero_PUJ_D_programa', 'Diplomado de Prueba', self.profile, self.config), ('Dip', 'filename'))
        with self.assertRaises(ValueError):
            g.degree_for('Ibero_PUJ_D_programa', 'Maestría de Prueba', self.profile, self.config)

    def test_multiple_sections_and_reject_malformed_html(self):
        paragraphs = g.word_paragraph('Diplomado de Prueba', 'title')
        for sub, text in [('Uno', 'primero.'), ('Dos', 'segundo')]:
            paragraphs += g.word_paragraph(sub, 'subtitle') + g.word_paragraph('<ul><li>' + text + '</li></ul>', 'code')
        record = g.extract(self.source(paragraphs))
        self.assertEqual([s['subtitle_original'] for s in record['sections']], ['Uno', 'Dos'])
        for bad in ['<ul><li>sin cierre</ul>', '<ul><li>padre<ul><li>hijo</li></ul></li></ul>']:
            with self.assertRaises(ValueError):
                parser = g.BulletHTML()
                parser.feed(bad)
                parser.result()

    def test_word_styles_and_copyable_html(self):
        html = g.render_html(self.base, ['Texto & contenido.'], self.profile)
        out = self.root / 'resultado.docx'
        g.make_docx(g.ASSETS / 'plantilla.docx', out, [{'title': 'Programa', 'sections': [{'subtitle': 'Competencias:', 'html': html}]}])
        with ZipFile(out) as z:
            root = ET.fromstring(z.read('word/document.xml'))
        paras = root.findall('.//w:body/w:p', g.NS)
        for i, color in [(1, 'yellow'), (2, 'green')]:
            rp = paras[i].find('w:r/w:rPr', g.NS)
            self.assertEqual(rp.find('w:sz', g.NS).get(g.q('val')), '24')
            self.assertEqual(rp.find('w:rFonts', g.NS).get(g.q('ascii')), 'Aptos')
            self.assertIsNotNone(rp.find('w:b', g.NS))
            self.assertEqual(rp.find('w:highlight', g.NS).get(g.q('val')), color)
        actual = []
        for p in paras[4:]:
            rp = p.find('w:r/w:rPr', g.NS)
            self.assertEqual(rp.find('w:sz', g.NS).get(g.q('val')), '18')
            self.assertEqual(rp.find('w:b', g.NS).get(g.q('val')), '0')
            self.assertEqual(rp.find('w:u', g.NS).get(g.q('val')), 'none')
            self.assertEqual(rp.find('w:highlight', g.NS).get(g.q('val')), 'none')
            actual.append(''.join(t.text or '' for t in p.findall('.//w:t', g.NS)))
        self.assertEqual('\n'.join(actual) + '\n', html)

    def test_delivery_contains_only_word_preview_and_validation(self):
        paragraphs = g.word_paragraph('Diplomado de Prueba', 'title')
        paragraphs += g.word_paragraph('Competencias', 'subtitle')
        paragraphs += g.word_paragraph('<ul><li>aprender.</li><li>crear sin punto</li></ul>', 'code')
        self.source(paragraphs)
        output = self.root / 'salida'
        with patch('sys.argv', ['generar.py', '--input', str(self.root), '--output', str(output)]):
            g.main()
        self.assertEqual({p.name for p in output.iterdir()},
                         {'IBERO_1_Dip.docx', 'revision.html', 'validacion.json'})
        report = json.loads((output / 'validacion.json').read_text())
        delivery = report['deliveries'][0]
        self.assertEqual((delivery['documents'], delivery['sections'], delivery['bullets']), (1, 1, 2))
        self.assertTrue(all(delivery['checks'].values()))
        self.assertIn('crear sin punto'.capitalize(), (output / 'revision.html').read_text())
        record = g.extract(self.root / 'Ibero_PUJ_D_prueba.docx')
        record['sections'][0].update(subtitle='Competencias', html=g.render_html(self.base, ['Texto incorrecto'], self.profile))
        with self.assertRaisesRegex(ValueError, 'HTML no conserva'):
            g.validate_delivery(g.ASSETS / 'plantilla.docx', output / 'IBERO_1_Dip.docx', [record], self.root)


if __name__ == '__main__':
    unittest.main()
