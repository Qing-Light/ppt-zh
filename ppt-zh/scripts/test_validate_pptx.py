"""Synthetic OOXML fixtures test checks, not PowerPoint rendering."""
import json
import tempfile
import unittest
import zipfile
from pathlib import Path
from validate_pptx import validate, NS

class ValidationTests(unittest.TestCase):
    def test_success_and_missing_conclusion(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            blocks = [{'id': name, 'page': page, 'source': 'source', 'evidence': 'evidence',
                       'draft': 'draft', 'humanized': 'result。', 'final': 'result。'}
                      for name, page in [('B1',3),('B2',3),('B3',3),('R1',4),('R2',4),('R3',4),('R4',4),('G',4),('F1',5),('L1',6),('L2',6),('L3',6)]]
            manifest = {'figures': [{'id':'Fig.1'}], 'blocks':blocks, 'pages':[{'page':i} for i in range(1,7)]}
            (d/'m.json').write_text(json.dumps(manifest))
            (d/'t.md').write_text('\n'.join('## '+b['id'] for b in blocks))
            (d/'p.png').write_bytes(b'preview placeholder')
            def build(omit=False):
                with zipfile.ZipFile(d/'test.pptx','w') as z:
                    z.writestr('ppt/presentation.xml', '<p:presentation xmlns:p="'+NS['p']+'" xmlns:r="'+NS['r']+'"><p:sldIdLst>'+''.join(f'<p:sldId id="{256+i}" r:id="r{i}"/>' for i in range(1,7))+'</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>')
                    z.writestr('ppt/_rels/presentation.xml.rels','<Relationships>'+''.join(f'<Relationship Id="r{i}" Target="slides/slide{i}.xml"/>' for i in range(1,7))+'</Relationships>')
                    for i in range(1,7):
                        names = ['block:'+b['id'] for b in blocks if b['page']==i and not (omit and b['id']=='F1')]
                        if i==1: names += ['presenter','date']
                        shapes = ''.join(f'<p:sp><p:nvSpPr><p:cNvPr name="{name}"/></p:nvSpPr><p:txBody><a:p><a:r><a:t>result。</a:t></a:r></a:p></p:txBody></p:sp>' for name in names)
                        if i==5: shapes += '<p:pic><p:nvPicPr><p:cNvPr name="figure:Fig.1"/></p:nvPicPr></p:pic>'
                        z.writestr(f'ppt/slides/slide{i}.xml',f'<p:sld xmlns:p="{NS["p"]}" xmlns:a="{NS["a"]}"><p:cSld><p:spTree>{shapes}</p:spTree></p:cSld></p:sld>')
            def run(): return validate(d/'test.pptx',d/'m.json',d/'t.md',d/'p.png')
            build()
            self.assertEqual(run()['status'], 'passed')
            build(True)
            self.assertTrue(any('F1: missing editable' in e for e in run()['errors']))
            build()
            (d/'t.md').write_text('')
            self.assertIn('Missing trace block headings',run()['errors'])
            (d/'p.png').unlink()
            self.assertIn('Preview missing or empty',run()['errors'])

if __name__ == '__main__': unittest.main()
