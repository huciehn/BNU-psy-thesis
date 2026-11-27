#!/usr/bin/env python3
"""Apply a reviewed paragraph/table plan to a new, content-preserving DOCX.

Only writes new files. Refuses unknown plan keys, uncovered content, stale plans,
unresolved revisions, and unsafe formatting-span edits. No Office installation
is required. Pagination and visual checks still require Word/WPS or a renderer.
"""
import argparse
import copy
import io
import json
import posixpath
import re
import shutil
import sys
import zipfile
from pathlib import Path
from lxml import etree
from docx_common import NS, REVISION_XPATH, qn, parse_xml, xml_bytes, sha256_file, read_docx, paragraphs, tables, owned_nodes, paragraph_text, has_content, fingerprint, load_profile

HEADER_REL = NS['r'] + '/header'
FOOTER_REL = NS['r'] + '/footer'

def element(parent, name, attrs=None):
    found = parent.find(name, NS)
    if found is None:
        found = etree.Element(qn(name))
        if name in ('w:tblPr', 'w:trPr', 'w:tcPr'):
            parent.insert(0, found)
        else:
            parent.append(found)
    if attrs:
        for k, v in attrs.items():
            found.set(qn(k) if ':' in k else k, str(v))
    return found

def remove(parent, names):
    for name in names:
        for n in list(parent.findall(name, NS)):
            parent.remove(n)

def ppr(p):
    prop = p.find('w:pPr', NS)
    if prop is None:
        prop = etree.Element(qn('w:pPr'))
        p.insert(0, prop)
    return prop

def rpr(r):
    prop = r.find('w:rPr', NS)
    if prop is None:
        prop = etree.Element(qn('w:rPr'))
        r.insert(0, prop)
    return prop

def boolprop(parent, tag, value):
    remove(parent, [tag])
    element(parent, tag, {'w:val': '1' if value else '0'})

def apply_run_properties(prop, rule):
    fonts = element(prop, 'w:rFonts')
    for name in ('asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'cstheme', 'csTheme'):
        fonts.attrib.pop(qn('w:' + name), None)
    for name in ('ascii', 'hAnsi', 'cs'):
        fonts.set(qn('w:' + name), rule['en_font'])
    fonts.set(qn('w:eastAsia'), rule['cn_font'])
    if rule['cn_font'] != rule['en_font']:
        fonts.set(qn('w:hint'), 'eastAsia')
    else:
        fonts.attrib.pop(qn('w:hint'), None)
    element(prop, 'w:sz', {'w:val': int(round(rule['size_pt'] * 2))})
    element(prop, 'w:szCs', {'w:val': int(round(rule['size_pt'] * 2))})
    if rule.get('font_color'):
        color = element(prop, 'w:color')
        color.attrib.clear()
        color.set(qn('w:val'), rule['font_color'])
    if 'bold' in rule:
        boolprop(prop, 'w:b', rule['bold'])
        boolprop(prop, 'w:bCs', rule['bold'])

def apply_paragraph_properties(prop, rule, safe_line=False):
    element(prop, 'w:jc', {'w:val': rule['align']})
    ind = element(prop, 'w:ind')
    # Clear twip overrides, including negative/right indentation, that defeat the profile.
    ind.attrib.clear()
    ind.set(qn('w:leftChars'), str(int(round(rule.get('left_chars', 0) * 100))))
    ind.set(qn('w:rightChars'), '0')
    if rule.get('hanging_chars', 0):
        ind.set(qn('w:hangingChars'), str(int(round(rule['hanging_chars'] * 100))))
    else:
        ind.set(qn('w:firstLineChars'), str(int(round(rule.get('first_line_chars', 0) * 100))))
    spacing = element(prop, 'w:spacing')
    spacing.attrib.clear()
    spacing.set(qn('w:before'), '0')
    spacing.set(qn('w:after'), '0')
    spacing.set(qn('w:beforeLines'), str(int(round(rule.get('before_lines', 0) * 100))))
    spacing.set(qn('w:afterLines'), str(int(round(rule.get('after_lines', 0) * 100))))
    if safe_line or rule.get('line_pt') is None:
        spacing.set(qn('w:line'), str(int(round(rule.get('line_multiple', 1) * 240))))
        spacing.set(qn('w:lineRule'), 'auto')
    else:
        spacing.set(qn('w:line'), str(int(round(rule['line_pt'] * 20))))
        spacing.set(qn('w:lineRule'), 'exact')
    boolprop(prop, 'w:keepNext', rule.get('keep_next', False))
    boolprop(prop, 'w:pageBreakBefore', rule.get('page_break_before', False))
    boolprop(prop, 'w:contextualSpacing', False)
    boolprop(prop, 'w:snapToGrid', False)
    boolprop(prop, 'w:widowControl', True)
    if 'outline_level' in rule:
        element(prop, 'w:outlineLvl', {'w:val': rule['outline_level']})
    else:
        # Direct outline levels from bad formatting can pollute the automatic TOC.
        element(prop, 'w:outlineLvl', {'w:val': 9})

def style_id(role):
    if re.fullmatch(r'heading[1-9]', role):
        return 'Heading' + role[-1]
    if re.fullmatch(r'toc[1-3]', role):
        return 'TOC' + role[-1]
    return 'BnuPsyThesis' + ''.join(s.capitalize() for s in role.split('_'))

def builtin_style_role(style):
    """Identify paragraph built-ins by semantic name, never by numeric IDs.

    WPS rewrites style IDs to numbers; localized Word documents can also use
    Chinese names. Updating every equivalent definition prevents a stale
    pre-existing built-in from winning when an application refreshes a TOC.
    """
    if style.get(qn('w:type')) != 'paragraph':
        return None
    labels = [style.get(qn('w:styleId'), '')]
    for node in style.findall('w:name', NS) + style.findall('w:aliases', NS):
        labels.extend(node.get(qn('w:val'), '').split(','))
    for label in labels:
        match = re.fullmatch(r'(?:toc|目录)\s*([1-3])', label.strip(), re.I)
        if match:
            return 'toc' + match.group(1)
        match = re.fullmatch(r'(?:heading|标题)\s*([1-9])', label.strip(), re.I)
        if match:
            return 'heading' + match.group(1)
    return None

def apply_toc_tabs(prop, profile):
    tabs = element(prop, 'w:tabs')
    remove(tabs, ['w:tab'])
    page = profile['page']
    usable = (page['width_mm'] - page['left_mm'] - page['right_mm']) * 1440 / 25.4
    etree.SubElement(tabs, qn('w:tab'), {qn('w:val'): 'right',
                    qn('w:leader'): profile.get('toc', {}).get('leader', 'middleDot'),
                    qn('w:pos'): str(round(usable))})

def ensure_styles(parts, profile):
    styles = parse_xml(parts['word/styles.xml']) if 'word/styles.xml' in parts else etree.Element(qn('w:styles'), nsmap={'w': NS['w']})
    if not styles.xpath('./w:style[@w:styleId="Normal"]', namespaces=NS):
        normal = etree.SubElement(styles, qn('w:style'), {qn('w:type'): 'paragraph', qn('w:styleId'): 'Normal', qn('w:default'): '1'})
        element(normal, 'w:name', {'w:val': 'Normal'})
    for role, overrides in profile['styles'].items():
        if overrides.get('preserve'):
            continue
        rule = dict(profile['defaults'], **overrides)
        sid = style_id(role)
        match = styles.xpath('./w:style[@w:styleId=$sid]', namespaces=NS, sid=sid)
        st = match[0] if match else etree.SubElement(styles, qn('w:style'), {qn('w:type'): 'paragraph', qn('w:styleId'): sid})
        builtin = bool(re.fullmatch(r'heading[1-9]|toc[1-3]', role))
        equivalent = [s for s in styles.findall('w:style', NS) if s is not st and builtin_style_role(s) == role] if builtin else []
        for definition in [st] + equivalent:
            if builtin:
                definition.attrib.pop(qn('w:customStyle'), None)
            else:
                definition.set(qn('w:customStyle'), '1')
            # Only the canonical definition needs its semantic name replaced;
            # preserve existing localized names referenced by native software.
            if definition is st:
                name = ('heading ' + role[-1]) if re.fullmatch(r'heading[1-9]', role) else ('toc ' + role[-1]) if re.fullmatch(r'toc[1-3]', role) else 'BNU-psy-thesis ' + role.replace('_', ' ')
                element(definition, 'w:name', {'w:val': name})
            element(definition, 'w:basedOn', {'w:val': 'Normal'})
            element(definition, 'w:next', {'w:val': style_id('body')})
            element(definition, 'w:qFormat')
            apply_paragraph_properties(element(definition, 'w:pPr'), rule)
            apply_run_properties(element(definition, 'w:rPr'), rule)
            if re.fullmatch(r'toc[1-3]', role):
                apply_toc_tabs(element(definition, 'w:pPr'), profile)
    parts['word/styles.xml'] = xml_bytes(styles)

def validate_plan(plan, original_hash, ps, ts, profile):
    allowed = {'profile', 'input_sha256', 'paragraphs', 'tables', 'text_edits'}
    extras = set(plan) - allowed
    if extras:
        raise ValueError('Unknown plan keys: %s' % ', '.join(sorted(extras)))
    if plan.get('profile') != profile['id']:
        raise ValueError('Plan profile must be %s' % profile['id'])
    if plan.get('input_sha256') != original_hash:
        raise ValueError('Stale plan: input_sha256 does not match this input DOCX')
    if plan.get('text_edits'):
        raise ValueError('text_edits is unsupported: this formatter guarantees unchanged content')
    mapping = plan.get('paragraphs')
    tmap = plan.get('tables')
    if not isinstance(mapping, dict) or not isinstance(tmap, dict):
        raise ValueError('Plan paragraphs and tables must be JSON objects')
    ids = {'p%05d' % i for i in range(len(ps))}
    if set(mapping) - ids:
        raise ValueError('Unknown paragraph IDs: %s' % ', '.join(sorted(set(mapping) - ids)))
    for i, p in enumerate(ps):
        key = 'p%05d' % i
        if key not in mapping:
            if has_content(p):
                raise ValueError('Missing reviewed role for nonempty paragraph %s: %s' % (key, paragraph_text(p)[:60]))
            continue
        spec = mapping[key]
        if not isinstance(spec, dict) or set(spec) - {'role', 'spans'}:
            raise ValueError('Paragraph %s accepts only role and spans' % key)
        if spec.get('role') not in profile['styles']:
            raise ValueError('Unknown role for %s: %r' % (key, spec.get('role')))
        spans = spec.get('spans', [])
        if not isinstance(spans, list):
            raise ValueError('spans must be a list for %s' % key)
        end_prev = 0
        for span in spans:
            if not isinstance(span, dict) or set(span) - {'start', 'end', 'italic', 'bold'}:
                raise ValueError('Invalid span keys for %s' % key)
            if type(span.get('start')) is not int or type(span.get('end')) is not int:
                raise ValueError('Span offsets must be integers for %s' % key)
            if not 0 <= span['start'] < span['end'] <= len(paragraph_text(p)) or span['start'] < end_prev:
                raise ValueError('Spans must be ordered, nonoverlapping, and within text for %s' % key)
            if not any(k in span for k in ('italic', 'bold')) or any(type(span[k]) is not bool for k in ('italic', 'bold') if k in span):
                raise ValueError('Span requires boolean italic and/or bold for %s' % key)
            end_prev = span['end']
    tids = {'t%05d' % i for i in range(len(ts))}
    if set(tmap) != tids:
        raise ValueError('Every table requires explicit kind; table IDs must exactly match inventory')
    for key, spec in tmap.items():
        if not isinstance(spec, dict) or set(spec) - {'kind', 'header_rows'}:
            raise ValueError('Table %s accepts only kind and header_rows' % key)
        if spec.get('kind') not in ('data', 'form', 'preserve'):
            raise ValueError('Table %s requires kind=data|form|preserve' % key)
        rows = ts[int(key[1:])].findall('w:tr', NS)
        n = spec.get('header_rows', 1)
        minimum = 1 if spec['kind'] == 'data' else 0
        if type(n) is not int or not minimum <= n <= max(1, len(rows)):
            raise ValueError('Invalid header_rows for %s' % key)

def split_plain_runs(p, cuts, key):
    """Split only plain direct runs; never flatten fields, links, or equations."""
    if not cuts:
        return
    allowed_children = {qn('w:pPr'), qn('w:r'), qn('w:bookmarkStart'), qn('w:bookmarkEnd'),
                        qn('w:commentRangeStart'), qn('w:commentRangeEnd'), qn('w:proofErr'),
                        qn('w:permStart'), qn('w:permEnd')}
    if any(c.tag not in allowed_children for c in p):
        raise ValueError('%s has complex content; spans/keyword label splits need a reviewed plain-text copy or specialized editor' % key)
    for r in p.findall('w:r', NS):
        if any(c.tag not in (qn('w:rPr'), qn('w:t'), qn('w:commentReference')) for c in r):
            raise ValueError('%s contains tabs, breaks, fields, drawings, or other complex runs; refusing unsafe split' % key)
    offset = 0
    for r in list(p.findall('w:r', NS)):
        text = ''.join(t.text or '' for t in r.findall('w:t', NS))
        lo, hi = offset, offset + len(text)
        inside = sorted(c for c in cuts if lo < c < hi)
        if inside:
            if r.find('w:commentReference', NS) is not None:
                raise ValueError('%s has a commentReference mixed inside a text run at the split boundary; refusing to move that marker' % key)
            index = p.index(r)
            points = [lo] + inside + [hi]
            for a, b in zip(points, points[1:]):
                new = etree.Element(qn('w:r'), dict(r.attrib))
                prop = r.find('w:rPr', NS)
                if prop is not None:
                    new.append(copy.deepcopy(prop))
                t = etree.SubElement(new, qn('w:t'))
                t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                t.text = text[a-lo:b-lo]
                p.insert(index, new)
                index += 1
            p.remove(r)
        offset = hi

def apply_spans(p, spans):
    offset = 0
    for r in p.findall('w:r', NS):
        length = sum(len(t.text or '') for t in r.findall('w:t', NS))
        lo, hi = offset, offset + length
        for span in spans:
            if lo >= span['start'] and hi <= span['end'] and length:
                prop = rpr(r)
                for key, tags in [('italic', ['w:i', 'w:iCs']), ('bold', ['w:b', 'w:bCs'])]:
                    if key in span:
                        for tag in tags:
                            boolprop(prop, tag, span[key])
        offset = hi

def ensure_blank_lines(root, ps, mapping):
    inserted = 0
    before_roles = {'abstract_label_cn', 'abstract_label_en', 'toc_title', 'keywords_cn', 'keywords_en'}
    after_roles = {'abstract_label_cn', 'abstract_label_en', 'toc_title'}
    for i, p in enumerate(ps):
        role = mapping.get('p%05d' % i, {}).get('role')
        if role not in before_roles | after_roles:
            continue
        parent = p.getparent()
        if parent is None or parent.tag != qn('w:body'):
            raise ValueError('Blank-line role %s must be a direct body paragraph' % role)
        def usable_blank(node):
            return (node is not None and node.tag == qn('w:p') and not has_content(node)
                    and node.find('w:pPr/w:sectPr', NS) is None
                    and not node.xpath('.//w:br[@w:type="page"]', namespaces=NS))
        def blank():
            node = etree.Element(qn('w:p'))
            element(ppr(node), 'w:pStyle', {'w:val': style_id('blank')})
            return node
        if role in before_roles and not usable_blank(p.getprevious()):
            p.addprevious(blank())
            inserted += 1
        if role in after_roles and not usable_blank(p.getnext()):
            p.addnext(blank())
            inserted += 1
    return inserted


def add_bookmarks(ps):
    ids = [int(n.get(qn('w:id'))) for p in ps for n in p.xpath('.//w:bookmarkStart | .//w:bookmarkEnd', namespaces=NS) if (n.get(qn('w:id')) or '').isdigit()]
    next_id = max(ids or [-1]) + 1
    names = {n.get(qn('w:name')) for p in ps for n in p.xpath('.//w:bookmarkStart', namespaces=NS)}
    if any(('bnu_psy_thesis_p%05d' % i) in names for i in range(len(ps))):
        raise ValueError('Input already contains BNU-psy-thesis mapping bookmarks; re-inventory a clean source or choose specialized reformatting')
    for i, p in enumerate(ps):
        start = etree.Element(qn('w:bookmarkStart'), {qn('w:id'): str(next_id), qn('w:name'): 'bnu_psy_thesis_p%05d' % i})
        end = etree.Element(qn('w:bookmarkEnd'), {qn('w:id'): str(next_id)})
        p.insert(1 if p.find('w:pPr', NS) is not None else 0, start)
        p.append(end)
        next_id += 1

def promote_inherited_formatting(p, source_styles, force_bold=False):
    """Keep semantic emphasis inherited from paragraph/character styles."""
    if source_styles is None:
        return
    styles = {s.get(qn('w:styleId')): s for s in source_styles.findall('w:style', NS)}
    defaults = source_styles.find('w:docDefaults/w:rPrDefault/w:rPr', NS)
    def chain(sid, seen=None):
        seen = seen or set()
        if not sid or sid in seen or sid not in styles:
            return []
        seen.add(sid)
        st = styles[sid]
        base = st.find('w:basedOn', NS)
        values = chain(base.get(qn('w:val')) if base is not None else None, seen)
        prop = st.find('w:rPr', NS)
        return values + ([prop] if prop is not None else [])
    ps = p.find('w:pPr/w:pStyle', NS)
    if ps is not None:
        sid = ps.get(qn('w:val'))
    else:
        sid = next((s.get(qn('w:styleId')) for s in styles.values() if s.get(qn('w:type')) == 'paragraph' and s.get(qn('w:default')) == '1'), 'Normal')
    props = ([defaults] if defaults is not None else []) + chain(sid)
    tags = ['w:i', 'w:iCs', 'w:vertAlign', 'w:u', 'w:strike', 'w:dstrike', 'w:caps', 'w:smallCaps']
    if not force_bold:
        tags += ['w:b', 'w:bCs']
    for r in owned_nodes(p, './/w:r'):
        direct = rpr(r)
        charstyle = direct.find('w:rStyle', NS)
        layers = props + chain(charstyle.get(qn('w:val')) if charstyle is not None else None)
        for tag in tags:
            if direct.find(tag, NS) is not None:
                continue
            resolved = next((pr.find(tag, NS) for pr in reversed(layers) if pr.find(tag, NS) is not None), None)
            if resolved is not None:
                direct.append(copy.deepcopy(resolved))

def format_paragraph(p, role, profile, key, spans=None, source_styles=None):
    rule = dict(profile['defaults'], **profile['styles'][role])
    if rule.get('preserve'):
        if spans:
            raise ValueError('preserve role does not accept spans: %s' % key)
        return 0
    spans = spans or []
    promote_inherited_formatting(p, source_styles, force_bold='bold' in rule)
    cuts = {s[c] for s in spans for c in ('start', 'end')}
    label_end = None
    if role in ('keywords_cn', 'keywords_en'):
        text = paragraph_text(p)
        match = re.match(r'^\s*关键词\s*[：:]', text) if role == 'keywords_cn' else re.match(r'^\s*KEY\s*WORDS\s*:', text, re.I)
        if not match:
            raise ValueError('%s keyword role lacks a recognizable label and colon; textual corrections require separate authorization' % key)
        label_end = match.end()
        # Whitespace after the English colon belongs to content.
        cuts.add(label_end)
    split_plain_runs(p, cuts, key)
    prop = ppr(p)
    element(prop, 'w:pStyle', {'w:val': style_id(role)})
    safe = bool(owned_nodes(p, './/w:drawing | .//w:pict | .//m:oMath | .//m:oMathPara | .//w:object'))
    apply_paragraph_properties(prop, rule, safe_line=safe)
    # Native TOC generation copies direct character formatting from source
    # headings. A 16 pt 黑体 run would therefore override the 12 pt 宋体 TOC
    # style. Keep heading font/size/bold in the paragraph style and retain only
    # semantic run emphasis (italics, superscripts, etc.) in the text itself.
    toc_source = 'outline_level' in rule and rule['outline_level'] <= 8
    if toc_source:
        mark_prop = prop.find('w:rPr', NS)
        if mark_prop is not None:
            remove(mark_prop, ['w:rFonts', 'w:sz', 'w:szCs', 'w:b', 'w:bCs', 'w:rStyle', 'w:color'])
        for r in owned_nodes(p, './/w:r'):
            run_prop = rpr(r)
            remove(run_prop, ['w:rFonts', 'w:sz', 'w:szCs', 'w:b', 'w:bCs', 'w:rStyle', 'w:color'])
    else:
        apply_run_properties(element(prop, 'w:rPr'), rule)
        for r in owned_nodes(p, './/w:r'):
            apply_run_properties(rpr(r), rule)
    if label_end is not None:
        offset = 0
        for r in p.findall('w:r', NS):
            text = ''.join(t.text or '' for t in r.findall('w:t', NS))
            label = offset < label_end
            kw_rule = dict(rule, bold=label)
            if role == 'keywords_cn':
                kw_rule['cn_font'] = '黑体' if label else '宋体'
            apply_run_properties(rpr(r), kw_rule)
            offset += len(text)
    apply_spans(p, spans)
    if role.startswith('toc') and role != 'toc_title':
        apply_toc_tabs(prop, profile)
    return fit_images(p, profile) if role == 'image' else 0


def fit_images(p, profile):
    resized = 0
    maximum = int(round(profile['page'].get('detector_max_image_mm', 150) * 36000))
    for container in owned_nodes(p, './/wp:inline | .//wp:anchor'):
        extent = container.find('wp:extent', NS)
        if extent is None:
            continue
        cx = int(extent.get('cx', '0'))
        cy = int(extent.get('cy', '0'))
        if cx <= maximum or cx <= 0 or cy <= 0:
            continue
        new_cx = maximum
        new_cy = round(cy * maximum / cx)
        extent.set('cx', str(new_cx))
        extent.set('cy', str(new_cy))
        for shape_extent in container.xpath('.//a:xfrm/a:ext', namespaces=NS):
            if shape_extent.get('cx') == str(cx) and shape_extent.get('cy') == str(cy):
                shape_extent.set('cx', str(new_cx))
                shape_extent.set('cy', str(new_cy))
        resized += 1
    return resized


def table_borders(tbl, spec, profile):
    if spec['kind'] != 'data':
        return
    prop = element(tbl, 'w:tblPr')
    element(prop, 'w:jc', {'w:val': 'center'})
    maximum = round(profile['page'].get('detector_max_table_mm', 150) * 1440 / 25.4)
    grid = tbl.findall('w:tblGrid/w:gridCol', NS)
    widths = [int(node.get(qn('w:w'), '0')) for node in grid]
    total = sum(widths)
    if total > maximum:
        if not grid or any(width <= 0 for width in widths):
            raise ValueError('Oversized data table lacks a usable grid; adjust its column widths in a reviewed copy')
        scale = maximum / total
        adjusted = [max(1, round(width * scale)) for width in widths]
        adjusted[-1] += maximum - sum(adjusted)
        for node, width in zip(grid, adjusted):
            node.set(qn('w:w'), str(width))
        element(prop, 'w:tblW', {'w:w': maximum, 'w:type': 'dxa'})
        element(prop, 'w:tblLayout', {'w:type': 'fixed'})
        remove(prop, ['w:tblInd'])
        for cell_width in tbl.xpath('.//w:tcPr/w:tcW[@w:type="dxa"]', namespaces=NS):
            old = int(cell_width.get(qn('w:w'), '0'))
            if old > 0:
                cell_width.set(qn('w:w'), str(max(1, round(old * scale))))
    # A table style can restore its own border conditionals unless direct borders
    # are explicitly nil; clear every direct cell border too.
    borders = element(prop, 'w:tblBorders')
    borders.clear()
    sizes = profile.get('tables', {})
    for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        sz = sizes.get('top_eighth_pt', 8) if side == 'top' else sizes.get('bottom_eighth_pt', 8) if side == 'bottom' else 0
        etree.SubElement(borders, qn('w:' + side), {qn('w:val'): 'single' if sz else 'nil', qn('w:sz'): str(sz), qn('w:color'): '000000'})
    rows = tbl.findall('w:tr', NS)
    headers = spec.get('header_rows', 1)
    for i, row in enumerate(rows):
        trprop = element(row, 'w:trPr')
        boolprop(trprop, 'w:cantSplit', True)
        boolprop(trprop, 'w:tblHeader', i < headers)
        for cell in row.findall('w:tc', NS):
            cellprop = element(cell, 'w:tcPr')
            cb = element(cellprop, 'w:tcBorders')
            cb.clear()
            for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV', 'tl2br', 'tr2bl'):
                sz = sizes.get('header_eighth_pt', 4) if side == 'bottom' and i == headers - 1 else 0
                # Explicit top/bottom at outer cells prevent old style borders
                # from competing with the three-line direct table border.
                if side == 'top' and i == 0:
                    sz = sizes.get('top_eighth_pt', 8)
                if side == 'bottom' and i == len(rows) - 1:
                    sz = sizes.get('bottom_eighth_pt', 8)
                etree.SubElement(cb, qn('w:' + side), {qn('w:val'): 'single' if sz else 'nil', qn('w:sz'): str(sz), qn('w:color'): '000000'})

def section_records(body):
    result = []
    start = 0
    children = list(body)
    for i, child in enumerate(children):
        sp = child.find('w:pPr/w:sectPr', NS) if child.tag == qn('w:p') else None
        if child.tag == qn('w:sectPr'):
            sp = child
        if sp is not None:
            result.append((start, i, sp))
            start = i + 1
    return result

def prepare_sections(root, ps, mapping, profile):
    body = root.find('w:body', NS)
    if body.find('w:sectPr', NS) is None:
        etree.SubElement(body, qn('w:sectPr'))
    odd = set(profile.get('requirements', {}).get('odd_section_roles', []))
    boundaries = []
    first_body = None
    first_appendix = None
    for i, p in enumerate(ps):
        role = mapping.get('p%05d' % i, {}).get('role')
        if role == 'heading1' and first_body is None:
            first_body = p
        if role in ('appendix_heading_single', 'appendix_heading_multi') and first_appendix is None:
            first_appendix = p
        if role in odd or p is first_body or p is first_appendix:
            if p.getparent() is not body:
                raise ValueError('Section-start role %s must be a direct body paragraph, not inside a table/text box/content control' % role)
            boundaries.append(p)
    for target in boundaries:
        idx = body.index(target)
        if idx == 0:
            continue
        current = next((record for record in section_records(body) if record[0] <= idx <= record[1]), None)
        if current is not None:
            leading = list(body)[current[0]:idx]
            reusable = all(
                node.tag in (qn('w:bookmarkStart'), qn('w:bookmarkEnd'))
                or (node.tag == qn('w:p') and not has_content(node)
                    and node.find('w:pPr/w:sectPr', NS) is None
                    and not node.xpath('.//w:br[@w:type="page"]', namespaces=NS)
                    and node.find('w:pPr/w:pageBreakBefore', NS) is None)
                for node in leading
            )
            if reusable:
                continue
        prev = body[idx-1]
        if prev.tag == qn('w:p') and prev.find('w:pPr/w:sectPr', NS) is not None:
            continue
        next_sect = next((sp for a, b, sp in section_records(body) if a <= idx <= b), None)
        if next_sect is None:
            raise ValueError('Could not locate section properties for boundary')
        new_sp = copy.deepcopy(next_sect)
        # Next section starts on its mandatory odd page. This preceding section
        # must retain its original start type, not inherit the next section's.
        if prev.tag != qn('w:p'):
            prev = etree.Element(qn('w:p'))
            body.insert(idx, prev)
        ppr(prev).append(new_sp)
    children = list(body)
    records = section_records(body)
    body_start_index = body.index(first_body) if first_body is not None else None
    targets = {body.index(p) for p in boundaries}
    first_body_section = None
    for k, (start, end, sp) in enumerate(records):
        if any(start <= i <= end for i in targets):
            element(sp, 'w:type', {'w:val': 'oddPage'})
        pg = element(sp, 'w:pgSz')
        page = profile['page']
        landscape = pg.get(qn('w:orient')) == 'landscape' or (int(pg.get(qn('w:w'), '0') or 0) > int(pg.get(qn('w:h'), '0') or 0) > 0)
        width, height = page['width_mm'], page['height_mm']
        if landscape:
            width, height = height, width
            pg.set(qn('w:orient'), 'landscape')
        else:
            pg.attrib.pop(qn('w:orient'), None)
        pg.set(qn('w:w'), str(round(width * 1440 / 25.4)))
        pg.set(qn('w:h'), str(round(height * 1440 / 25.4)))
        margin = element(sp, 'w:pgMar')
        for side in ('top', 'bottom', 'left', 'right'):
            margin.set(qn('w:' + side), str(round(page[side + '_mm'] * 1440 / 25.4)))
        margin.set(qn('w:gutter'), '0')
        if body_start_index is not None and start <= body_start_index <= end:
            first_body_section = k
        number = element(sp, 'w:pgNumType')
        number.attrib.clear()
        number.set(qn('w:fmt'), 'decimal')
        if k == first_body_section:
            number.set(qn('w:start'), str(profile['page'].get('body_page_start', 1)))
    return records, first_body_section

def clean_footer(root, name):
    """Remove existing PAGE fields only where all other content is preserved."""
    for simple in list(root.xpath('.//w:fldSimple', namespaces=NS)):
        if re.match(r'^\s*PAGE(?:\s|$)', simple.get(qn('w:instr'), ''), re.I):
            simple.getparent().remove(simple)
    for p in list(root.xpath('.//w:p', namespaces=NS)):
        instr = ''.join(p.xpath('./w:r/w:instrText/text()', namespaces=NS))
        if not re.search(r'\bPAGE\b', instr, re.I):
            continue
        # Complex fields are safe only when the PAGE field is the complete
        # paragraph. Mixing labels or another field needs a specialized editor.
        direct_runs = p.findall('w:r', NS)
        if any(c.tag not in (qn('w:pPr'), qn('w:r')) for c in p):
            raise ValueError('Unsafe mixed PAGE footer in %s; preserve/edit this footer manually' % name)
        in_field = False
        outside_text = ''
        begin_count = 0
        for r in direct_runs:
            for n in r:
                if n.tag == qn('w:fldChar'):
                    kind = n.get(qn('w:fldCharType'))
                    if kind == 'begin':
                        begin_count += 1
                        in_field = True
                    elif kind == 'end':
                        in_field = False
                elif n.tag == qn('w:t') and not in_field:
                    outside_text += n.text or ''
                elif n.tag not in (qn('w:rPr'), qn('w:t'), qn('w:instrText')):
                    raise ValueError('Unsafe complex PAGE footer object in %s' % name)
        if outside_text.strip() or begin_count != 1 or in_field or not re.match(r'^\s*PAGE(?:\s|$)', instr, re.I):
            raise ValueError('PAGE footer mixed with other text/fields in %s; refusing content loss' % name)
        p.getparent().remove(p)
    return root

def add_page_field(footer, profile):
    page = profile['page']
    p = etree.SubElement(footer, qn('w:p'))
    prop = ppr(p)
    element(prop, 'w:jc', {'w:val': 'center'})
    element(prop, 'w:spacing', {'w:before': 0, 'w:after': 0, 'w:line': 240, 'w:lineRule': 'auto'})
    rule = {'cn_font': page.get('footer_font', 'Times New Roman'), 'en_font': page.get('footer_font', 'Times New Roman'), 'size_pt': page.get('footer_size_pt', 9), 'bold': False}
    apply_run_properties(element(prop, 'w:rPr'), rule)
    for kind in ('begin', 'code', 'separate', 'result', 'end'):
        r = etree.SubElement(p, qn('w:r'))
        apply_run_properties(rpr(r), rule)
        if kind == 'code':
            n = etree.SubElement(r, qn('w:instrText'))
            n.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            n.text = ' PAGE '
        elif kind == 'result':
            etree.SubElement(r, qn('w:t')).text = '1'
        else:
            etree.SubElement(r, qn('w:fldChar'), {qn('w:fldCharType'): kind})

def configure_footers(parts, records, first_body_section, profile):
    relname = 'word/_rels/document.xml.rels'
    rels = parse_xml(parts[relname]) if relname in parts else etree.Element(qn('pr:Relationships'), nsmap={None: NS['pr']})
    content_types = parse_xml(parts['[Content_Types].xml'])
    rel_by_id = {r.get('Id'): r for r in rels}
    next_id = max([int(r.get('Id')[3:]) for r in rels if re.fullmatch(r'rId\d+', r.get('Id', ''))] or [0]) + 1
    footer_no = 1
    while 'word/bnuPsyThesisFooter%d.xml' % footer_no in parts:
        footer_no += 1
    inherited = {}
    cache = {}
    for k, (_, _, sp) in enumerate(records):
        for ref in sp.findall('w:footerReference', NS):
            inherited[ref.get(qn('w:type'), 'default')] = ref.get(qn('r:id'))
        numbered = first_body_section is not None and k >= first_body_section
        original_refs = dict(inherited)
        remove(sp, ['w:footerReference'])
        for kind in ('default', 'even', 'first'):
            rid = original_refs.get(kind) or original_refs.get('default')
            cachekey = (rid, numbered)
            if cachekey in cache:
                new_id = cache[cachekey]
            else:
                if rid:
                    rel = rel_by_id.get(rid)
                    if rel is None or rel.get('Type') != FOOTER_REL or rel.get('TargetMode') == 'External':
                        raise ValueError('Invalid/external footer relationship %r' % rid)
                    target = rel.get('Target', '')
                    name = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('word', target))
                    if name not in parts:
                        raise ValueError('Missing footer part %s' % name)
                    # Copy related parts untouched. Our new footer remains in
                    # word/, so relationship targets keep their meaning.
                    footer = clean_footer(parse_xml(parts[name]), name)
                    old_relpath = posixpath.join(posixpath.dirname(name), '_rels', posixpath.basename(name) + '.rels')
                else:
                    footer = etree.Element(qn('w:ftr'), nsmap={'w': NS['w'], 'r': NS['r']})
                    old_relpath = None
                if numbered:
                    add_page_field(footer, profile)
                if len(footer) == 0:
                    # Empty unnumbered sections need no footer reference.
                    new_id = None
                else:
                    if not footer.findall('w:p', NS) and not footer.findall('w:tbl', NS):
                        etree.SubElement(footer, qn('w:p'))
                    name_new = 'word/bnuPsyThesisFooter%d.xml' % footer_no
                    footer_no += 1
                    new_id = 'rId%d' % next_id
                    next_id += 1
                    parts[name_new] = xml_bytes(footer)
                    if old_relpath and old_relpath in parts:
                        new_relpath = 'word/_rels/' + posixpath.basename(name_new) + '.rels'
                        parts[new_relpath] = parts[old_relpath]
                    etree.SubElement(rels, qn('pr:Relationship'), {'Id': new_id, 'Type': FOOTER_REL, 'Target': posixpath.basename(name_new)})
                    etree.SubElement(content_types, qn('ct:Override'), {'PartName': '/' + name_new, 'ContentType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'})
                cache[cachekey] = new_id
            if new_id:
                # sectPr schema order requires header/footer references first.
                sp.insert(0, etree.Element(qn('w:footerReference'), {qn('w:type'): kind, qn('r:id'): new_id}))
    parts[relname] = xml_bytes(rels)
    parts['[Content_Types].xml'] = xml_bytes(content_types)

def register_standard_parts(parts):
    """Minimal DOCX inputs can omit styles/settings and their relationships."""
    relname = 'word/_rels/document.xml.rels'
    rels = parse_xml(parts[relname])
    ct = parse_xml(parts['[Content_Types].xml'])
    next_id = max([int(r.get('Id')[3:]) for r in rels if re.fullmatch(r'rId\d+', r.get('Id', ''))] or [0]) + 1
    for short in ('styles', 'settings'):
        if not any(r.get('Type', '').endswith('/' + short) for r in rels):
            etree.SubElement(rels, qn('pr:Relationship'), {'Id': 'rId%d' % next_id, 'Type': NS['r'] + '/' + short, 'Target': short + '.xml'})
            next_id += 1
        partname = '/word/' + short + '.xml'
        if not any(n.get('PartName') == partname for n in ct):
            etree.SubElement(ct, qn('ct:Override'), {'PartName': partname, 'ContentType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.' + short + '+xml'})
    parts[relname] = xml_bytes(rels)
    parts['[Content_Types].xml'] = xml_bytes(ct)

def sort_properties(root):
    # WordprocessingML property containers have ordered child sequences.
    orders = {
        qn('w:pPr'): 'pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange',
        qn('w:rPr'): 'rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath rPrChange',
        qn('w:sectPr'): 'headerReference footerReference footnotePr endnotePr type pgSz pgMar paperSrc pgBorders lnNumType pgNumType cols formProt vAlign noEndnote titlePg textDirection bidi rtlGutter docGrid printerSettings sectPrChange',
        qn('w:tblPr'): 'tblStyle tblpPr tblOverlap bidiVisual tblStyleRowBandSize tblStyleColBandSize tblW jc tblCellSpacing tblInd tblBorders shd tblLayout tblCellMar tblLook tblCaption tblDescription tblPrChange',
        qn('w:trPr'): 'cnfStyle divId gridBefore gridAfter wBefore wAfter cantSplit trHeight tblHeader tblCellSpacing jc hidden ins del trPrChange',
        qn('w:tcPr'): 'cnfStyle tcW gridSpan hMerge vMerge tcBorders shd noWrap tcMar textDirection tcFitText vAlign hideMark headers cellIns cellDel cellMerge tcPrChange',
    }
    for tag, sequence in orders.items():
        rank = {qn('w:' + name): i for i, name in enumerate(sequence.split())}
        for prop in root.iter(tag):
            if any(a.tag in (qn('m:oMath'), qn('m:oMathPara')) for a in prop.iterancestors()):
                continue
            ordered = sorted(list(prop), key=lambda n: rank.get(n.tag, len(rank)))
            prop[:] = ordered

def format_document(input_path, plan_path, profile_path=None):
    original_hash = sha256_file(input_path)
    source_parts = read_docx(input_path)
    if sha256_file(input_path) != original_hash:
        raise ValueError('Input changed during reading; regenerate inventory from a stable copy')
    parts = dict(source_parts)
    root = parse_xml(parts['word/document.xml'])
    if root.xpath(REVISION_XPATH, namespaces=NS):
        raise ValueError('Unresolved tracked changes: accept/reject changes in a reviewed copy, then re-inventory')
    profile = load_profile(profile_path)
    with open(plan_path, encoding='utf-8-sig') as f:
        plan = json.load(f)
    ps, ts = paragraphs(root), tables(root)
    validate_plan(plan, original_hash, ps, ts, profile)
    before = fingerprint(source_parts)
    source_styles = parse_xml(source_parts['word/styles.xml']) if 'word/styles.xml' in source_parts else None
    ensure_styles(parts, profile)
    effective_roles = {}
    resized_images = 0
    for i, p in enumerate(ps):
        key = 'p%05d' % i
        if key not in plan['paragraphs']:
            continue
        spec = plan['paragraphs'][key]
        role = spec['role']
        # Form tables use their independently specified line spacing; they must
        # not acquire a three-line research-table format.
        ancestor_table = next((a for a in p.iterancestors() if a.tag == qn('w:tbl')), None)
        if ancestor_table is not None:
            index = ts.index(ancestor_table)
            kind = plan['tables']['t%05d' % index]['kind']
            if kind == 'form' and role == 'table_cell':
                role = 'form_cell'
        resized_images += format_paragraph(p, role, profile, key, spec.get('spans'), source_styles)
        effective_roles[key] = role
    add_bookmarks(ps)
    inserted_blanks = ensure_blank_lines(root, ps, plan['paragraphs'])
    for i, tbl in enumerate(ts):
        table_borders(tbl, plan['tables']['t%05d' % i], profile)
    records, first_body_section = prepare_sections(root, ps, plan['paragraphs'], profile)
    configure_footers(parts, records, first_body_section, profile)
    settings = parse_xml(parts['word/settings.xml']) if 'word/settings.xml' in parts else etree.Element(qn('w:settings'), nsmap={'w': NS['w']})
    boolprop(settings, 'w:doNotCompressPictures', True)
    boolprop(settings, 'w:updateFields', True)
    parts['word/settings.xml'] = xml_bytes(settings)
    register_standard_parts(parts)
    for name in list(parts):
        if name.startswith('word/bnuPsyThesisFooter') and name.endswith('.xml'):
            footer = parse_xml(parts[name])
            sort_properties(footer)
            parts[name] = xml_bytes(footer)
    for name in ('word/styles.xml',):
        st = parse_xml(parts[name])
        sort_properties(st)
        parts[name] = xml_bytes(st)
    sort_properties(root)
    parts['word/document.xml'] = xml_bytes(root)
    after = fingerprint(parts)
    if before != after:
        mismatches = [key for key in before if before[key] != after.get(key)]
        raise ValueError('Content-preservation verification failed: %s' % ', '.join(mismatches))
    changed = sorted(name for name in parts if source_parts.get(name) != parts[name])
    untouched = sorted(name for name in source_parts if name not in changed)
    for name in untouched:
        if parts[name] != source_parts[name]:
            raise ValueError('Unexpected change to protected ZIP part: ' + name)
    if sha256_file(input_path) != original_hash:
        raise ValueError('Input changed during formatting; regenerate inventory and plan from the updated copy')
    manual = ['all_pages_visual', 'physical_odd_pages', 'toc_page_accuracy', 'fonts_available', 'official_cover',
              'citations_bijective', 'reference_metadata_and_italics', 'statistics_preserved', 'figures_and_tables']
    instructions = ''.join(root.xpath('.//w:instrText/text() | .//w:fldSimple/@w:instr', namespaces=NS))
    if not re.search(r'\bTOC\b', instructions, re.I):
        manual.append('missing_automatic_toc')
    for name in ('word/footnotes.xml', 'word/endnotes.xml'):
        if name in parts:
            manual.append(name + '_formatting')
    for name, data in source_parts.items():
        if name.startswith('word/header') and name.endswith('.xml'):
            header = parse_xml(data)
            header_instr = ''.join(header.xpath('.//w:instrText/text() | .//w:fldSimple/@w:instr', namespaces=NS))
            if re.search(r'\bPAGE\b', header_instr, re.I):
                manual.append('header_page_fields:' + name)
    if first_body_section is None:
        manual.append('missing_body_heading1_page_number_start')
    report = {'schema_version': 1, 'profile': profile['id'], 'profile_version': profile.get('version'),
              'input_sha256': original_hash, 'plan_sha256': sha256_file(plan_path),
              'before': before, 'after': after, 'content_preserved': True,
              'paragraph_count': len(ps), 'table_count': len(ts), 'section_count': len(records),
              'inserted_blank_paragraphs': inserted_blanks, 'resized_images': resized_images,
              'effective_roles': effective_roles, 'changed_parts': changed, 'unchanged_parts': untouched,
              'unchanged_parts_byte_identical': True, 'manual_required': manual,
              'status': 'formatted_pending_render_and_review',
              'page_note': 'OOXML sets oddPage starts and Arabic page numbering. Actual pages and TOC results must be refreshed and rendered in Word/WPS.'}
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)
    return buf.getvalue(), report

def preserve_unformatted(input_path, output_path):
    input_path = Path(input_path)
    output_path = Path(output_path)
    base = output_path.parent / (input_path.stem + '_未排版')
    candidate = base.with_suffix('.docx')
    number = 2
    while candidate.exists():
        candidate = base.with_name(base.name + '_' + str(number)).with_suffix('.docx')
        number += 1
    candidate.parent.mkdir(parents=True, exist_ok=True)
    with input_path.open('rb') as source, candidate.open('xb') as target:
        shutil.copyfileobj(source, target)
    shutil.copystat(input_path, candidate)
    return candidate


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input', type=Path)
    parser.add_argument('--plan', required=True, type=Path)
    parser.add_argument('--output', required=True, type=Path)
    parser.add_argument('--report', required=True, type=Path)
    parser.add_argument('--profile', type=Path)
    args = parser.parse_args()
    try:
        paths = [args.input.resolve(), args.plan.resolve(), args.output.resolve(), args.report.resolve()]
        if len(set(paths)) != len(paths):
            raise ValueError('Input, plan, output, and report paths must all differ')
        for path in (args.output, args.report):
            if path.exists():
                raise ValueError('Refusing to overwrite existing output: %s' % path)
        preserved = preserve_unformatted(args.input, args.output)
        data, report = format_document(args.input, args.plan, args.profile)
        report['unformatted_copy'] = str(preserved)
        report['output_sha256'] = __import__('hashlib').sha256(data).hexdigest()
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.report.parent.mkdir(parents=True, exist_ok=True)
        with args.output.open('xb') as f:
            f.write(data)
        with args.report.open('x', encoding='utf-8') as f:
            json.dump(report, f, ensure_ascii=False, indent=2)
            f.write('\n')
        print('Formatted DOCX: %s\nUnformatted original: %s\nVerification report: %s' % (args.output, preserved, args.report))
    except Exception as e:
        print('ERROR: %s' % e, file=sys.stderr)
        return 2
    return 0

if __name__ == '__main__':
    sys.exit(main())
