"""Shared, dependency-light OOXML helpers for BNU-psy-thesis (Python 3.9+)."""
import hashlib
import json
import posixpath
import re
import zipfile
from pathlib import Path
from lxml import etree

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'pr': 'http://schemas.openxmlformats.org/package/2006/relationships',
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
}

REVISION_XPATH = './/w:ins | .//w:del | .//w:moveFrom | .//w:moveTo | .//w:rPrChange | .//w:pPrChange | .//w:tblPrChange | .//w:trPrChange | .//w:tcPrChange | .//w:sectPrChange | .//w:tblGridChange | .//w:moveFromRangeStart | .//w:moveToRangeStart'

def qn(name):
    prefix, local = name.split(':', 1)
    return '{%s}%s' % (NS[prefix], local)

def parse_xml(data):
    return etree.fromstring(data, etree.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False))

def xml_bytes(root):
    return etree.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True)

def sha256_file(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for block in iter(lambda: f.read(1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()

def read_docx(path):
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        if len(names) != len(set(names)):
            raise ValueError('DOCX contains duplicate ZIP entries')
        if 'word/document.xml' not in names:
            raise ValueError('Input is not a Word DOCX package')
        if sum(i.file_size for i in z.infolist()) > 512 * 1024 * 1024:
            raise ValueError('DOCX exceeds the 512 MiB uncompressed safety limit')
        return {name: z.read(name) for name in names}

def paragraphs(root):
    body = root.find('w:body', NS)
    return body.xpath('.//w:p', namespaces=NS) if body is not None else []

def tables(root):
    body = root.find('w:body', NS)
    return body.xpath('.//w:tbl', namespaces=NS) if body is not None else []

def owned_nodes(p, xpath):
    """Exclude text-box paragraphs nested inside a drawing in this paragraph."""
    found = []
    for node in p.xpath(xpath, namespaces=NS):
        nearest = next((a for a in node.iterancestors() if a.tag == qn('w:p')), None)
        if nearest is p:
            found.append(node)
    return found

def paragraph_text(p):
    out = []
    for n in owned_nodes(p, './/w:t | .//m:t | .//w:r/w:tab | .//w:r/w:br | .//w:r/w:cr'):
        if n.tag in (qn('w:t'), qn('m:t')):
            out.append(n.text or '')
        elif n.tag == qn('w:tab'):
            out.append('\t')
        elif n.get(qn('w:type')) != 'page':
            out.append('\n')
    return ''.join(out)

def has_content(p):
    return bool(paragraph_text(p).strip() or owned_nodes(p, './/w:drawing | .//w:pict | .//m:oMath | .//m:oMathPara | .//w:sym | .//w:instrText | .//w:fldSimple | .//w:object'))

def _digest(value):
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, separators=(',', ':'), sort_keys=True).encode('utf-8')).hexdigest()

def fingerprint(parts):
    root = parse_xml(parts['word/document.xml'])
    # A formatting operation may split a text run or add an empty section break.
    # Compare whole paragraph text, including equation text, excluding empty breaks.
    text = [paragraph_text(p) for p in paragraphs(root) if paragraph_text(p)]
    formulas = [etree.tostring(n, method='c14n').decode('utf-8') for n in root.xpath('.//m:oMath[not(ancestor::m:oMath)]', namespaces=NS)]
    protected = []
    for n in root.xpath('.//w:sym | .//w:object | .//w:footnoteReference | .//w:endnoteReference | .//w:fldChar | .//w:instrText', namespaces=NS):
        protected.append(etree.tostring(n, method='c14n').decode('utf-8'))
    for n in root.xpath('.//w:fldSimple', namespaces=NS):
        protected.append(['simple_field', dict(n.attrib)])
    media = {n: hashlib.sha256(b).hexdigest() for n, b in parts.items() if n.startswith('word/media/')}
    image_rels = set()
    for name, data in sorted(parts.items()):
        if name.endswith('.rels'):
            rels = parse_xml(data)
            for r in rels:
                if r.get('Type', '').endswith('/image'):
                    # Copying a footer with its image relationships creates
                    # equivalent references; compare distinct definitions with
                    # their resolution base rather than the new footer name.
                    image_rels.add(json.dumps([posixpath.dirname(posixpath.dirname(name)), dict(r.attrib)], sort_keys=True))
    return {'text_sha256': _digest(text), 'formula_sha256': _digest(formulas),
            'protected_objects_sha256': _digest(protected),
            'image_parts_sha256': _digest(media), 'image_relationships_sha256': _digest(sorted(image_rels)),
            'nonempty_text_paragraphs': len(text), 'formula_count': len(formulas), 'media_count': len(media)}

def load_profile(path=None):
    path = Path(path) if path else Path(__file__).resolve().parent.parent / 'references' / 'profile.json'
    with path.open(encoding='utf-8-sig') as f:
        return json.load(f)

def suggested_role(p):
    text = paragraph_text(p).strip()
    style = p.find('w:pPr/w:pStyle', NS)
    style_id = style.get(qn('w:val'), '') if style is not None else ''
    if not text:
        if owned_nodes(p, './/m:oMath | .//m:oMathPara'):
            return 'formula', True
        if owned_nodes(p, './/w:drawing | .//w:pict | .//w:object'):
            return 'image', True
        return None, False
    if any(a.tag == qn('w:tc') for a in p.iterancestors()):
        return 'table_cell', True
    match = re.search(r'(?:heading|标题)\s*([1-9])$', style_id, re.I)
    if match:
        return 'heading' + match.group(1), True
    match = re.search(r'toc\s*([1-3])$', style_id, re.I)
    if match:
        return 'toc' + match.group(1), True
    compact = re.sub(r'\s+', '', text)
    fixed = {'摘要': 'abstract_label_cn', 'ABSTRACT': 'abstract_label_en', '目录': 'toc_title', 'Contents': 'toc_title',
             '参考文献': 'reference_heading', 'References': 'reference_heading', '致谢': 'ack_heading'}
    if compact in fixed:
        return fixed[compact], True
    if re.match(r'^关键词\s*[：:]', text):
        return 'keywords_cn', True
    if re.match(r'^KEY\s*WORDS\s*:', text, re.I):
        return 'keywords_en', True
    if re.match(r'^表\s*\d+\s+', text):
        return 'table_caption', True
    if re.match(r'^图\s*\d+\s+', text):
        return 'figure_caption', True
    if compact.startswith('附录'):
        return 'appendix_heading_single', True
    if re.match(r'^\d+(?:\.\d+){0,8}\s+\S', text) and len(text) < 100:
        levels = text.split()[0].count('.') + 1
        return 'heading' + str(levels), True
    return 'body', True
