#!/usr/bin/env python3
"""Inventory an existing DOCX without changing its content or formatting."""
import argparse
import json
import sys
from pathlib import Path
from docx_common import NS, REVISION_XPATH, qn, read_docx, parse_xml, sha256_file, paragraphs, tables, paragraph_text, suggested_role, has_content, fingerprint

def inventory(path):
    parts = read_docx(path)
    root = parse_xml(parts['word/document.xml'])
    p_items = []
    for i, p in enumerate(paragraphs(root)):
        role, ambiguous = suggested_role(p)
        style = p.find('w:pPr/w:pStyle', NS)
        p_items.append({'id': 'p%05d' % i, 'text': paragraph_text(p),
                        'style': style.get(qn('w:val')) if style is not None else None,
                        'suggested_role': role, 'ambiguous': ambiguous, 'requires_role': has_content(p),
                        'in_table': any(a.tag == qn('w:tc') for a in p.iterancestors()),
                        'has_image': bool(p.xpath('.//w:drawing | .//w:pict', namespaces=NS)),
                        'has_formula': bool(p.xpath('.//m:oMath | .//m:oMathPara', namespaces=NS))})
    t_items = []
    for i, t in enumerate(tables(root)):
        rows = t.findall('w:tr', NS)
        t_items.append({'id': 't%05d' % i, 'rows': len(rows), 'max_cells': max([len(r.findall('w:tc', NS)) for r in rows] or [0]),
                        'sample': [[paragraph_text(p) for p in c.xpath('.//w:p', namespaces=NS)] for r in rows[:2] for c in r.findall('w:tc', NS)],
                        'suggested_kind': None, 'ambiguous': True})
    warnings = []
    for name in ('word/footnotes.xml', 'word/endnotes.xml', 'word/comments.xml'):
        if name in parts:
            warnings.append('%s is preserved; inspect its formatting separately' % name)
    if root.xpath(REVISION_XPATH, namespaces=NS):
        warnings.append('Unresolved tracked changes must be accepted/rejected in a reviewed copy before formatting')
    return {'schema_version': 1, 'input_sha256': sha256_file(path), 'paragraphs': p_items, 'tables': t_items,
            'fingerprint': fingerprint(parts), 'warnings': warnings,
            'instruction': 'Suggestions are intentionally ambiguous. Review every nonempty paragraph and every table before creating a plan.'}

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input', type=Path)
    parser.add_argument('--out', required=True, type=Path)
    args = parser.parse_args()
    try:
        if args.out.resolve() == args.input.resolve():
            raise ValueError('Inventory output must not replace input')
        result = inventory(args.input)
        args.out.parent.mkdir(parents=True, exist_ok=True)
        with args.out.open('x', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
            f.write('\n')
        print('Inventory written: %s' % args.out)
    except Exception as e:
        print('ERROR: %s' % e, file=sys.stderr)
        return 2
    return 0

if __name__ == '__main__':
    sys.exit(main())
