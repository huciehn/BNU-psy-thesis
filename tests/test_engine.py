"""Behavioral tests use isolated artifacts under the skill's _local directory."""
import hashlib
import json
import os
import sys
import unittest
import uuid
import zipfile
import struct
import zlib
from pathlib import Path
from types import SimpleNamespace
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm
from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
from docx_common import NS, paragraph_text, paragraphs, read_docx, sha256_file
from format_docx import format_document
from validate_docx import validate
from build_docx import build
import update_skill
from update_skill import version_tuple
from site_check import validate_credentials
from make_cover import extract_logo, make_blocks


def create_fixture(folder, full=False, form=False):
    d = Document()
    role_by_text = {}

    def p(text, role):
        para = d.add_paragraph(text)
        role_by_text[text] = role
        return para

    if full:
        p("论文格式示例", "cover_title")
        p("姓名：示例作者", "cover_meta")
        p("2026年10月", "cover_date")
        p("论文格式示例中文题目", "abstract_title_cn")
        p("摘  要", "abstract_label_cn")
        p("本段仅为排版验证用文本，不代表真实研究。", "abstract_body_cn")
        p("关键词：排版，验证，心理学", "keywords_cn")
        p("A Formatting Example", "abstract_title_en")
        p("ABSTRACT", "abstract_label_en")
        p("This text is a formatting fixture for testing document preservation.", "abstract_body_en")
        p("KEY WORDS: Formatting, Validation, Psychology", "keywords_en")
        p("目  录", "toc_title")
        toc = d.add_paragraph()
        role_by_text[""] = "toc1"
        r = toc.add_run()
        begin = OxmlElement("w:fldChar"); begin.set(qn("w:fldCharType"), "begin")
        ins = OxmlElement("w:instrText"); ins.text = ' TOC \\o "1-3" \\h \\z \\u '
        separate = OxmlElement("w:fldChar"); separate.set(qn("w:fldCharType"), "separate")
        end = OxmlElement("w:fldChar"); end.set(qn("w:fldCharType"), "end")
        for node in (begin, ins, separate, end): r._r.append(node)
    p("1  前言", "heading1")
    body = p("需要保留原文及统计结果：M = 12.34，SD = 3.45。", "body")
    body.clear()
    body.add_run("需要保留原文及统计结果：")
    body.add_run("M").italic = True
    body.add_run(" = 12.34，")
    body.add_run("SD").italic = True
    body.add_run(" = 3.45。")
    # Symbol runs are invisible to python-docx paragraph.text; they must survive.
    sym = OxmlElement("w:sym"); sym.set(qn("w:font"), "Symbol"); sym.set(qn("w:char"), "00B1")
    body.add_run()._r.append(sym)
    p("1.1 研究方法", "heading2")
    p("普通正文含中文、English，以及数字2026。", "body")
    math = d.add_paragraph()
    om = OxmlElement("m:oMath")
    mr = OxmlElement("m:r"); mt = OxmlElement("m:t"); mt.text = "x²"
    mr.append(mt); om.append(mr); math._p.append(om)
    role_by_text["x²"] = "formula"
    image_path = folder / "fixture.png"
    def chunk(kind, data):
        return struct.pack('!I', len(data)) + kind + data + struct.pack('!I', zlib.crc32(kind + data) & 0xffffffff)
    image_path.write_bytes(b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('!2I5B',160,100,8,2,0,0,0))
                           + chunk(b'IDAT', zlib.compress((b'\x00' + b'\xff' * 480) * 100)) + chunk(b'IEND', b''))
    image = d.add_paragraph(); image.add_run().add_picture(str(image_path), width=Cm(2))
    p("图1 排版测试图", "figure_caption")
    p("表1 格式测试数据", "table_caption")
    table = d.add_table(rows=3, cols=2)
    for row, values in zip(table.rows, (("变量", "结果"), ("均值", "12.34"), ("标准差", "3.45"))):
        for cell, value in zip(row.cells, values): cell.text = value
    p("注：* p < 0.05，** p < 0.01。", "table_note")
    p("参考文献", "reference_heading")
    ref_text = "示例作者. (2021). 排版演示. 示例期刊, 12(1), 1-2."
    p(ref_text, "reference")
    if full:
        p("附录  A", "appendix_heading_single")
        p("本附录仅用于测试格式。", "appendix_body")
        p("致  谢", "ack_heading")
        p("感谢参与排版测试的同学。", "ack_body")
    path = folder / "input.docx"; d.save(path)
    parts = read_docx(path); root = etree.fromstring(parts["word/document.xml"])
    plan = {"profile": "bnu-psychology-2021", "input_sha256": sha256_file(path), "paragraphs": {}, "tables": {"t00000": {"kind": "form" if form else "data", "header_rows": 1}}}
    for i, para in enumerate(paragraphs(root)):
        text = paragraph_text(para)
        inside = next(para.iterancestors(qn("w:tc")), None) is not None
        role = "table_cell" if inside else role_by_text.get(text)
        if para.findall(".//w:drawing", NS): role = "image"
        if role:
            plan["paragraphs"]["p{:05d}".format(i)] = {"role": role}
        if role == 'table_note':
            import re
            plan['paragraphs']['p{:05d}'.format(i)]['spans'] = [
                {'start': m.start(), 'end': m.end(), 'italic': True}
                for m in re.finditer(r'p(?=\s*<)', text)]
    rid = next(pid for pid, item in plan["paragraphs"].items() if item["role"] == "reference")
    start = ref_text.index("示例期刊")
    plan["paragraphs"][rid]["spans"] = [{"start": start, "end": start + len("示例期刊, 12"), "italic": True}]
    pp = folder / "plan.json"; pp.write_text(json.dumps(plan, ensure_ascii=False), encoding="utf-8")
    return path, pp, plan


class EngineTests(unittest.TestCase):
    def setUp(self):
        self.folder = ROOT / "_local/tests" / uuid.uuid4().hex
        self.folder.mkdir(parents=True)

    def formatted(self, full=False, form=False):
        src, plan_path, plan = create_fixture(self.folder, full, form)
        data, changes = format_document(src, plan_path)
        out = self.folder / "formatted.docx"; out.write_bytes(data)
        return src, out, plan, changes

    def test_text_math_image_symbol_and_significance_stars_survive(self):
        src, out, plan, changes = self.formatted()
        self.assertTrue(changes["content_preserved"])
        self.assertTrue(changes["unchanged_parts_byte_identical"])
        report = validate(out, plan, src, mode="partial")
        self.assertTrue(report["structural_pass"], report["errors"])
        self.assertFalse(report["ready_to_submit"])
        self.assertIn("** p < 0.01", Document(out).paragraphs[-3].text)

    def test_complete_thesis_structure_passes_but_requires_render(self):
        src, out, plan, changes = self.formatted(full=True)
        report = validate(out, plan, src)
        self.assertTrue(report["structural_pass"], report["errors"])
        self.assertFalse(report["ready_to_submit"])
        self.assertIn("physical_odd_pages", report["pending_reviews"])

    def test_form_is_not_converted_to_research_table(self):
        src, out, plan, changes = self.formatted(form=True)
        report = validate(out, plan, src, mode="partial")
        self.assertTrue(report["structural_pass"], report["errors"])
        self.assertIn("form_cell", changes["effective_roles"].values())

    def test_stale_plan_and_missing_roles_fail_before_write(self):
        src, pp, plan = create_fixture(self.folder)
        before = sha256_file(src)
        plan["input_sha256"] = "0" * 64
        pp.write_text(json.dumps(plan), encoding="utf-8")
        with self.assertRaises(ValueError): format_document(src, pp)
        self.assertEqual(before, sha256_file(src))
        plan["input_sha256"] = before
        plan["paragraphs"].pop(next(iter(plan["paragraphs"])))
        pp.write_text(json.dumps(plan), encoding="utf-8")
        with self.assertRaises(ValueError): format_document(src, pp)

    def test_independent_checker_catches_corrupted_margin_and_text(self):
        src, out, plan, _ = self.formatted()
        parts = read_docx(out); root = etree.fromstring(parts["word/document.xml"])
        root.find(".//w:pgMar", NS).set(qn("w:top"), "1")
        root.find(".//w:t", NS).text = "被错误覆盖的原文"
        parts["word/document.xml"] = etree.tostring(root)
        damaged = self.folder / "damaged.docx"
        with zipfile.ZipFile(damaged, "w") as z:
            for name, b in parts.items(): z.writestr(name, b)
        report = validate(damaged, plan, src, mode="partial")
        rules = {e["rule"] for e in report["errors"]}
        self.assertIn("MARGIN", rules)
        self.assertIn("PRESERVATION", rules)

    def test_old_review_cannot_approve_changed_document(self):
        src, out, plan, _ = self.formatted(full=True)
        review = {"docx_sha256": "0" * 64, "rendered_pdf_sha256": "1" * 64, "reviewer": "test", "checks": {}}
        report = validate(out, plan, src, review)
        self.assertFalse(report["ready_to_submit"])
        self.assertIn("REVIEW_BINDING", {e["rule"] for e in report["errors"]})

    def test_checker_rejects_nonitalic_statistical_variable(self):
        src, out, plan, _ = self.formatted()
        parts = read_docx(out); root = etree.fromstring(parts["word/document.xml"])
        run = next(r for r in root.findall('.//w:r', NS) if ''.join(r.xpath('./w:t/text()', namespaces=NS)) == 'M')
        run.find('w:rPr/w:i', NS).set(qn('w:val'), '0')
        parts['word/document.xml'] = etree.tostring(root)
        damaged = self.folder / 'upright-statistic.docx'
        with zipfile.ZipFile(damaged, 'w') as z:
            for name, data in parts.items(): z.writestr(name, data)
        report = validate(damaged, plan, src, mode='partial')
        self.assertIn('STAT_ITALICS', {e['rule'] for e in report['errors']})

    def test_structured_builder_uses_same_formatter_and_refuses_overwrite(self):
        content = self.folder / 'content.json'
        content.write_text(json.dumps({'blocks': [
            {'role': 'heading1', 'text': '1  Introduction'},
            {'role': 'body', 'text': 'Verified content for a partial formatting test.'}]}), encoding='utf-8')
        src, pp = self.folder / 'built.docx', self.folder / 'built-plan.json'
        build(content, src, pp)
        data, _ = format_document(src, pp)
        out = self.folder / 'built-formatted.docx'; out.write_bytes(data)
        report = validate(out, json.loads(pp.read_text(encoding='utf-8')), src, mode='partial')
        self.assertTrue(report['structural_pass'], report['errors'])
        with self.assertRaises(ValueError): build(content, src, pp)

    def test_native_toc_cache_can_replace_its_internal_bookmark(self):
        src, out, plan, _ = self.formatted(full=True)
        parts = read_docx(out); root = etree.fromstring(parts['word/document.xml'])
        toc = next(p for p in paragraphs(root) if p.find('w:pPr/w:pStyle', NS) is not None
                   and p.find('w:pPr/w:pStyle', NS).get(qn('w:val')) == 'TOC1')
        for mark in list(toc.findall('w:bookmarkStart', NS)) + list(toc.findall('w:bookmarkEnd', NS)):
            toc.remove(mark)
        run = toc.find('w:r', NS)
        end = run.find('w:fldChar[@w:fldCharType="end"]', NS)
        text = etree.Element(qn('w:t')); text.text = '1  前言'
        run.insert(run.index(end), text)
        parts['word/document.xml'] = etree.tostring(root)
        native = self.folder / 'toc-updated.docx'
        with zipfile.ZipFile(native, 'w') as z:
            for name, data in parts.items(): z.writestr(name, data)
        report = validate(native, plan, src)
        self.assertTrue(report['structural_pass'], report['errors'])
        self.assertTrue(report['checks']['content_preserved'])

    def test_site_reported_format_rules_are_checked(self):
        src, out, plan, _ = self.formatted(full=True)
        parts = read_docx(out)
        root = etree.fromstring(parts['word/document.xml'])
        for text in ('摘  要', 'ABSTRACT', '目  录'):
            para = next(p for p in paragraphs(root) if paragraph_text(p) == text)
            spacing = para.find('w:pPr/w:spacing', NS)
            spacing.set(qn('w:beforeLines'), '100')
            spacing.set(qn('w:afterLines'), '100')
        for text in ('参考文献', '注：* p < 0.05，** p < 0.01。'):
            para = next(p for p in paragraphs(root) if paragraph_text(p) == text)
            para.find('w:pPr/w:jc', NS).set(qn('w:val'), 'left')
        root.find('.//w:tblPr/w:jc', NS).set(qn('w:val'), 'left')
        caption = next(p for p in paragraphs(root) if paragraph_text(p) == '图1 排版测试图')
        caption.find('.//w:t', NS).text = '图1　排版测试图'
        for grid in root.findall('.//w:tblGrid/w:gridCol', NS):
            grid.set(qn('w:w'), '5000')
        extent = root.find('.//wp:extent', NS)
        extent.set('cx', str(165 * 36000))
        parts['word/document.xml'] = etree.tostring(root)
        damaged = self.folder / 'damaged-site-rules.docx'
        with zipfile.ZipFile(damaged, 'w') as z:
            for name, data in parts.items():
                z.writestr(name, data)
        report = validate(damaged, plan, src)
        rules = {e['rule'] for e in report['errors']}
        self.assertIn('SPACE', rules)
        self.assertIn('ALIGN', rules)
        self.assertIn('TABLE_ALIGN', rules)
        self.assertIn('TABLE_WIDTH', rules)
        self.assertIn('IMAGE_WIDTH', rules)
        self.assertIn('CAPTION_NUMBER', rules)

    def test_oversized_images_and_tables_are_scaled_without_changing_content(self):
        src, pp, plan = create_fixture(self.folder)
        parts = read_docx(src)
        root = etree.fromstring(parts['word/document.xml'])
        extent = root.find('.//wp:extent', NS)
        old_cy = int(extent.get('cy'))
        extent.set('cx', str(165 * 36000))
        extent.set('cy', str(round(old_cy * 1.5)))
        for grid in root.findall('.//w:tblGrid/w:gridCol', NS):
            grid.set(qn('w:w'), '4800')
        parts['word/document.xml'] = etree.tostring(root)
        with zipfile.ZipFile(src, 'w') as z:
            for name, data in parts.items():
                z.writestr(name, data)
        plan['input_sha256'] = sha256_file(src)
        pp.write_text(json.dumps(plan), encoding='utf-8')
        data, changes = format_document(src, pp)
        out = self.folder / 'scaled.docx'
        out.write_bytes(data)
        self.assertEqual(changes['resized_images'], 1)
        self.assertTrue(changes['content_preserved'])
        report = validate(out, plan, src, mode='partial')
        self.assertTrue(report['structural_pass'], report['errors'])

    def test_detector_cover_and_appendix_roles(self):
        content = self.folder / 'cover-content.json'
        content.write_text(json.dumps({'blocks': [
            {'role': 'cover_main_title', 'text': '本科生毕业论文(设计)'},
            {'role': 'cover_label', 'text': '毕业论文（设计）题目：'},
            {'role': 'cover_title', 'text': '双语儿童感数能力的发展'},
            {'role': 'cover_meta', 'text': '部   院  系：心理学部'},
            {'role': 'blank', 'text': ''},
            {'role': 'cover_date', 'text': '2026年11月25日'},
            {'role': 'appendix_heading_single', 'text': '附录  A'},
            {'role': 'appendix_body', 'text': '真实补充材料说明。'}
        ]}), encoding='utf-8')
        src, pp = self.folder / 'cover-source.docx', self.folder / 'cover-plan.json'
        build(content, src, pp)
        data, _ = format_document(src, pp)
        out = self.folder / 'cover-formatted.docx'
        out.write_bytes(data)
        report = validate(out, json.loads(pp.read_text(encoding='utf-8')), src, mode='partial')
        self.assertTrue(report['structural_pass'], report['errors'])

    def test_update_and_site_input_validation(self):
        self.assertGreater(version_tuple('v2.3.0'), version_tuple('2.2.0'))
        with self.assertRaises(ValueError):
            version_tuple('latest')
        validate_credentials('202421161111', 'test@example.edu.cn')
        with self.assertRaises(ValueError):
            validate_credentials('123', 'test@example.edu.cn')
        with self.assertRaises(ValueError):
            validate_credentials('202421161111', 'invalid')

    def test_cover_uses_user_supplied_template_logo(self):
        src, _, _ = create_fixture(self.folder)
        logo = self.folder / 'extracted-logo.png'
        extract_logo(src, logo)
        self.assertEqual(logo.read_bytes(), (self.folder / 'fixture.png').read_bytes())
        args = SimpleNamespace(title='测试论文', student_id='202421161111', student_name='测试',
                               advisor='测试', department='心理学部', major='心理学', advisor_title='',
                               advisor_unit='北京师范大学心理学部', date='2026年11月25日')
        blocks = make_blocks(args, logo.name)
        self.assertEqual(blocks[0]['path'], logo.name)
        self.assertEqual(blocks[1]['role'], 'cover_main_title')
        self.assertEqual(blocks[-1]['role'], 'cover_date')
        self.assertEqual(blocks[-2]['role'], 'blank')

    def test_update_install_works_from_inside_skill_directory(self):
        old_root = self.folder / 'installed'
        staged_parent = self.folder / 'staging'
        staged = staged_parent / 'BNU-psy-thesis'
        old_root.mkdir()
        (staged / 'tests').mkdir(parents=True)
        (staged / 'tests' / 'test_minimal.py').write_text('import unittest\n\nclass MinimalTest(unittest.TestCase):\n    def test_release(self):\n        self.assertTrue(True)\n', encoding='utf-8')
        (old_root / 'SKILL.md').write_text('old', encoding='utf-8')
        (staged / 'SKILL.md').write_text('---\nname: BNU-psy-thesis\n---\n', encoding='utf-8')
        (staged / 'scripts').mkdir()
        (staged / 'scripts' / 'new_script.py').write_text('VALUE = 1\n', encoding='utf-8')
        (staged_parent / 'BNU-psy-thesis.skill').write_bytes(b'new-release')
        (staged_parent / 'manifest.json').write_text('{}', encoding='utf-8')
        (old_root / 'dist').mkdir()
        (old_root / 'dist' / 'old.skill').write_bytes(b'old-release')
        previous_root = update_skill.ROOT
        previous_cwd = Path.cwd()
        try:
            update_skill.ROOT = old_root
            os.chdir(old_root)
            backup = update_skill.install(staged_parent, 'BNU-psy-thesis')
        finally:
            os.chdir(previous_cwd)
            update_skill.ROOT = previous_root
        self.assertEqual((old_root / 'SKILL.md').read_text(encoding='utf-8'), '---\nname: BNU-psy-thesis\n---\n')
        self.assertTrue((old_root / 'scripts' / 'new_script.py').is_file())
        self.assertEqual((backup / 'SKILL.md').read_text(encoding='utf-8'), 'old')
        self.assertEqual((old_root / 'dist' / 'BNU-psy-thesis.skill').read_bytes(), b'new-release')
        self.assertFalse((old_root / 'dist' / 'old.skill').exists())
        self.assertEqual((backup / 'dist' / 'old.skill').read_bytes(), b'old-release')


if __name__ == "__main__":
    unittest.main()
