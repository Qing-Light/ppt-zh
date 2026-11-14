#!/usr/bin/env python3
"""Dependency-free structural checks. Does not render or validate scientific claims."""
import argparse
import json
import re
import sys
import zipfile
import posixpath
from pathlib import Path
import xml.etree.ElementTree as ET

NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}

def validate(pptx, manifest, trace, preview):
    errors, warnings = [], []
    def check(ok, message):
        if not ok:
            errors.append(message)
    m = json.loads(Path(manifest).read_text(encoding='utf-8'))
    figures = m['figures']
    n = len(figures)
    check(n > 0, 'No figures in manifest')
    check(len({f['id'] for f in figures}) == n, 'Duplicate figure IDs')
    blocks = m['blocks']
    ids = [b['id'] for b in blocks]
    routes = [i for i in ids if re.fullmatch(r'R[1-6]', i)]
    expected = {f'B{i}' for i in range(1, 4)} | {'G'} | {f'L{i}' for i in range(1, 4)} | {f'F{i}' for i in range(1, n+1)}
    check(4 <= len(routes) <= 6, 'Expected 4–6 route blocks')
    expected |= {f'R{i}' for i in range(1, len(routes)+1)}
    check(set(ids) == expected and len(ids) == len(set(ids)), 'Incomplete or duplicate block IDs')
    check(len(m['pages']) == n+5, 'Manifest page count mismatch')
    check([p['page'] for p in m['pages']] == list(range(1, n+6)), 'Manifest page order mismatch')
    trace_text = Path(trace).read_text(encoding='utf-8')
    headings = set(re.findall(r'^##\s+([BRFL]\d+|G)\s*$', trace_text, re.M))
    check(expected <= headings, 'Missing trace block headings')
    check(Path(preview).is_file() and Path(preview).stat().st_size > 0, 'Preview missing or empty')
    with zipfile.ZipFile(pptx) as z:
        check(z.testzip() is None, 'ZIP CRC failure')
        presentation = ET.fromstring(z.read('ppt/presentation.xml'))
        rels = ET.fromstring(z.read('ppt/_rels/presentation.xml.rels'))
        targets = {r.attrib['Id']: r.attrib['Target'] for r in rels}
        slide_ids = presentation.findall('p:sldIdLst/p:sldId', NS)
        check(len(slide_ids) == n+5, 'PPTX page count mismatch')
        size = presentation.find('p:sldSz', NS)
        width, height = int(size.attrib['cx']), int(size.attrib['cy'])
        slides = []
        for num, sid in enumerate(slide_ids, 1):
            target = targets[sid.attrib['{'+NS['r']+'}id']]
            path = target.lstrip('/') if target.startswith('/') else posixpath.normpath('ppt/'+target)
            root = ET.fromstring(z.read(path))
            objects = {}
            for obj in list(root.findall('.//p:sp', NS)) + list(root.findall('.//p:graphicFrame', NS)) + list(root.findall('.//p:pic', NS)):
                prop = obj.find('.//p:cNvPr', NS)
                if prop is not None:
                    objects[prop.attrib.get('name', '')] = obj
                off, ext = obj.find('.//a:xfrm/a:off', NS), obj.find('.//a:xfrm/a:ext', NS)
                if off is not None and ext is not None:
                    x, y = int(off.attrib['x']), int(off.attrib['y'])
                    w, h = int(ext.attrib['cx']), int(ext.attrib['cy'])
                    check(x >= 0 and y >= 0 and x+w <= width+9144 and y+h <= height+9144,
                          f'Page {num}: object outside slide bounds')
                    if obj.tag.endswith('pic') and w*h >= width*height*.95:
                        warnings.append(f'Page {num}: near-full-page image; inspect editability')
            slides.append(objects)
        def text(obj):
            return ''.join(t.text or '' for t in obj.findall('.//a:t', NS))
        for name in ('presenter', 'date'):
            obj = slides[0].get(name) if slides else None
            check(obj is not None and bool(text(obj)), f'Missing editable cover field: {name}')
        for b in blocks:
            page = b.get('page', 0)
            check(isinstance(page, int) and 1 <= page <= len(slides), f'{b["id"]}: invalid page')
            if not isinstance(page, int) or not 1 <= page <= len(slides):
                continue
            for key in ('source', 'evidence', 'draft', 'humanized', 'final'):
                check(bool(b.get(key)), f'{b["id"]}: missing {key}')
            obj = slides[page-1].get('block:'+b['id'])
            if obj is None and b['id'].startswith('R'):
                obj = slides[page-1].get('methods-table')
            check(obj is not None and bool(text(obj)), f'{b["id"]}: missing editable object')
            if obj is not None:
                check(bool(b.get('final')) and b['final'] in text(obj), f'{b["id"]}: final text mismatch')
            if b['id'].startswith('F'):
                check(b.get('final', '').count('。') == 1 and b.get('final', '').endswith('。'), f'{b["id"]}: expected one concluding full stop')
        for i, f in enumerate(figures):
            page = i+5
            objects = slides[page-1] if page <= len(slides) else {}
            prefix = 'figure:'+f['id']
            check(any((name == prefix or name.startswith(prefix+':')) and obj.tag.endswith('pic') for name, obj in objects.items()), f'Page {page}: missing figure {f["id"]}')
    warnings.append('Rendering, line count, fonts, clipping, image fidelity and scientific accuracy require separate review.')
    return {'status': 'failed' if errors else 'passed', 'scope': 'structural_only', 'errors': errors, 'warnings': warnings}

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('pptx')
    for arg in ('manifest', 'trace', 'preview'):
        parser.add_argument('--'+arg, required=True)
    args = parser.parse_args()
    try:
        result = validate(args.pptx, args.manifest, args.trace, args.preview)
    except (OSError, ValueError, KeyError, TypeError, ET.ParseError, zipfile.BadZipFile) as exc:
        result = {'status': 'failed', 'scope': 'structural_only', 'errors': [str(exc)], 'warnings': []}
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 1 if result['status'] == 'failed' else 0

if __name__ == '__main__':
    sys.exit(main())
