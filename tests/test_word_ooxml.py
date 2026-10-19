"""Read-side Word normalization, unit construction, and textconv tests."""

import hashlib
import importlib.util
import io
import json
import os
import re
from pathlib import Path
import subprocess
import sys
import unittest
import warnings
import zipfile

from word_fixtures import (
    MAIN_TYPES, PNG, PR, R, STRICT_R, STRICT_W, W, default_styles, hyperlink,
    image, make_docx, numbering_xml, p, relationships, run, specification_docx,
    table, tracked, unpack, xml_document, zip_parts,
)


ROOT = Path(__file__).resolve().parents[1]
MODULE = ROOT / 'scripts' / 'word_ooxml.py'


class WordTestCase(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.module_name = '_word_ooxml_read_tests'
        spec = importlib.util.spec_from_file_location(cls.module_name, MODULE)
        cls.word = importlib.util.module_from_spec(spec)
        sys.modules[cls.module_name] = cls.word
        spec.loader.exec_module(cls.word)

    @classmethod
    def tearDownClass(cls):
        sys.modules.pop(cls.module_name, None)

    def document(self, body='', **kwargs):
        return self.word.load(make_docx(body, **kwargs))

    def lines(self, data):
        return tuple(line for story in self.word.load(data).stories for unit in story.units for line in unit.lines)

    def first_line(self, body='', **kwargs):
        return self.document(body, **kwargs).stories[0].units[0].lines[0]

    def assert_noise(self, before, after):
        self.assertNotEqual(before, after)
        self.assertEqual(self.lines(before), self.lines(after))
        self.assertEqual(self.word.render_text(before), self.word.render_text(after))


class NormalizationTests(WordTestCase):
    def test_run_splits_empty_runs_and_xml_space_are_noise(self):
        before = make_docx(p('合成の段落'))
        after = make_docx(p(runs='<w:r><w:rPr><w:noProof/></w:rPr><w:t>合成の</w:t></w:r>'
                                '<w:r/><w:r><w:t xml:space="default"> 段落 </w:t></w:r>'))
        self.assert_noise(before, after)

    def test_word_save_attributes_and_proofing_are_noise(self):
        before = make_docx(p('保存の比較'))
        body = ('<w:p w:rsidR="01020304" w:rsidRPr="05060708" w14:paraId="11223344" w14:textId="55667788">'
                '<w:pPr><w:cnfStyle w:val="000000000000"/></w:pPr>'
                '<w:proofErr w:type="spellStart"/><w:r><w:rPr><w:lang w:val="ja-JP" w:eastAsia="ja-JP"/>'
                '<w:rFonts w:hint="eastAsia"/><w:noProof/></w:rPr><w:t>保存の比較</w:t>'
                '<w:lastRenderedPageBreak/></w:r><w:proofErr w:type="spellEnd"/></w:p>')
        self.assert_noise(before, make_docx(body))

    def test_hidden_bookmark_names_ids_and_spans_are_noise(self):
        before = make_docx(p('段落甲') + p('段落乙'))
        for name in ('_GoBack', '_TocSynthetic', '_HlkSynthetic', '_RefSynthetic'):
            with self.subTest(name=name):
                body = ('<w:bookmarkStart w:id="99" w:name="' + name + '"/>'
                        + p('段落甲') + p('段落乙') + '<w:bookmarkEnd w:id="99"/>')
                self.assert_noise(before, make_docx(body))

    def test_revision_ids_and_modern_utc_timestamp_are_noise(self):
        before = p(runs=tracked('追記した文字', ident=3))
        after = p(runs=tracked('追記した文字', ident=71).replace('w:author=', 'w16du:dateUtc="2026-09-08T01:02:03Z" w:author='))
        self.assert_noise(make_docx(before), make_docx(after))

    def test_drawing_names_ids_anchor_ids_and_creation_ids_are_noise(self):
        before = image(doc_id=4, name='Picture old')
        after = image(doc_id=987, name='Picture new').replace('<wp:inline>', '<wp:inline wp14:anchorId="ABCDEF12" wp14:editId="1234ABCD">')
        after = after.replace('id="0" name="Synthetic pixel"', 'id="87" name="Different pixel"')
        after = after.replace('<a:blip r:embed="rImage"/>', '<a:blip r:embed="rImage"><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId id="arbitrary"/></a:ext></a:extLst></a:blip>')
        kwargs = {'rels': [('rImage', 'image', 'media/pixel.png')], 'parts': {'word/media/pixel.png': PNG}}
        self.assert_noise(make_docx(p(runs=before), **kwargs), make_docx(p(runs=after), **kwargs))

    def test_table_look_val_and_explicit_flags_are_equivalent(self):
        old = '<w:tblLook w:val="04A0"/>'
        new = '<w:tblLook w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
        self.assert_noise(make_docx(table([['試験']], props=old) + p()),
                          make_docx(table([['試験']], props=new) + p()))

    def test_japanese_and_english_style_identifiers_use_the_same_names(self):
        japanese = p('見出しの合成例', style='1') + p(runs=run('本文', '<w:rStyle w:val="a0"/>'))
        english = p('見出しの合成例', style='Heading1') + p(runs=run('本文', '<w:rStyle w:val="DefaultParagraphFont"/>'))
        self.assert_noise(make_docx(japanese), make_docx(english, japanese_ids=False))

    def test_list_ids_and_definition_metadata_are_noise(self):
        def item(ident):
            return p('項目', props='<w:numPr><w:ilvl w:val="0"/><w:numId w:val="' + str(ident) + '"/></w:numPr>')
        before = make_docx(item(7), numbering=numbering_xml())
        changed = numbering_xml(27, 15, 'FFFFFFFF', 'EEEEEEEE').replace('A1B2C3D4', '87654321').replace('w:tentative="1"', 'w:tentative="0"')
        self.assert_noise(before, make_docx(item(27), numbering=changed))

    def test_note_and_comment_identifiers_are_replaced_by_content(self):
        for kind in ('footnote', 'endnote', 'comment'):
            with self.subTest(kind=kind):
                extra = ' w:author="合成者" w:date="2026-02-03T04:05:06Z"' if kind == 'comment' else ''
                kwargs_before = {kind + 's': '<w:' + kind + ' w:id="9"' + extra + '>' + p('同じ注記') + '</w:' + kind + '>'}
                kwargs_after = {kind + 's': '<w:' + kind + ' w:id="18"' + extra + '>' + p('同じ注記') + '</w:' + kind + '>'}
                before = make_docx(p(runs=run('注記') + '<w:r><w:' + kind + 'Reference w:id="9"/></w:r>'), **kwargs_before)
                after = make_docx(p(runs=run('注記') + '<w:r><w:' + kind + 'Reference w:id="18"/></w:r>'), **kwargs_after)
                self.assert_noise(before, after)

    def test_run_font_attribute_order_and_run_splitting_are_noise(self):
        first = '<w:rFonts w:ascii="Aptos" w:hAnsi="Aptos" w:eastAsia="游明朝"/>'
        reordered = '<w:rFonts w:eastAsia="游明朝" w:ascii="Aptos" w:hAnsi="Aptos"/>'
        self.assert_noise(make_docx(p(runs=run('甲乙', first))),
                          make_docx(p(runs=run('甲', reordered) + run('乙', first))))

    def test_relationship_ids_and_image_part_names_are_noise(self):
        before = make_docx(p(runs=image('rBefore') + hyperlink('同じリンク', rel_id='rLinkBefore')),
                           rels=[('rBefore', 'image', 'media/old.png'), ('rLinkBefore', 'hyperlink', 'https://example.invalid/same', 'External')],
                           parts={'word/media/old.png': PNG})
        after = make_docx(p(runs=image('rAfter') + hyperlink('同じリンク', rel_id='rLinkAfter')),
                          rels=[('rAfter', 'image', 'media/new.png'), ('rLinkAfter', 'hyperlink', 'https://example.invalid/same', 'External')],
                          parts={'word/media/new.png': PNG})
        self.assert_noise(before, after)

    def test_numbering_level_style_references_use_names(self):
        number = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr>'
        japanese = numbering_xml().replace('<w:lvlJc', '<w:pStyle w:val="1"/><w:lvlJc')
        english = numbering_xml().replace('<w:lvlJc', '<w:pStyle w:val="Heading1"/><w:lvlJc')
        self.assert_noise(make_docx(p('番号項目', props=number), numbering=japanese),
                          make_docx(p('番号項目', props=number), numbering=english, japanese_ids=False))

    def test_plain_paragraph_has_no_hash_or_direct_format(self):
        for body in (p('文字だけ'), p('明示スタイル', style='1'), p()):
            with self.subTest(body=body):
                self.assertEqual(self.first_line(body).format, '')


class SensitivityTests(WordTestCase):
    def test_direct_run_formatting_changes_hash_without_changing_text(self):
        base = self.first_line(p('甲乙'))
        properties = {
            'bold': '<w:b/>', 'color': '<w:color w:val="1379BD"/>',
            'size': '<w:sz w:val="27"/>', 'highlight': '<w:highlight w:val="yellow"/>',
            'hidden': '<w:vanish/>', 'superscript': '<w:vertAlign w:val="superscript"/>',
            'spacing': '<w:spacing w:val="30"/>', 'explicit-off': '<w:b w:val="0"/>',
        }
        for name, prop in properties.items():
            with self.subTest(name=name):
                changed = self.first_line(p(runs=run('甲', prop) + run('乙')))
                self.assertEqual(changed.text, base.text)
                self.assertNotEqual(changed.format, base.format)
                self.assertRegex(changed.format, r'#[0-9a-f]{8}$')

    def test_direct_paragraph_formatting_changes_hash_without_changing_text(self):
        base = self.first_line(p('試験する'))
        for prop in ('<w:jc w:val="center"/>', '<w:ind w:left="240"/>', '<w:spacing w:after="120"/>'):
            with self.subTest(prop=prop):
                changed = self.first_line(p('試験する', props=prop))
                self.assertEqual(changed.text, base.text)
                self.assertNotEqual(changed.format, base.format)

    def test_explicit_style_changes_record_label(self):
        before = self.first_line(p('同じ文字列', style='1'))
        after = self.first_line(p('同じ文字列', style='2'))
        self.assertEqual(before.text, after.text)
        self.assertEqual((before.label, after.label), ('heading 1', 'heading 2'))
        self.assertNotEqual(before.key, after.key)

    def test_list_level_changes_hash_and_uses_definition_not_computed_number(self):
        lines = []
        for level in (0, 1):
            props = '<w:numPr><w:ilvl w:val="' + str(level) + '"/><w:numId w:val="7"/></w:numPr>'
            lines.append(self.first_line(p('一覧の項目', props=props), numbering=numbering_xml()))
        self.assertEqual(lines[0].text, lines[1].text)
        self.assertIn('番号「%1.」L1', lines[0].format)
        self.assertIn('番号「%2.」L2', lines[1].format)
        self.assertNotEqual(lines[0].format, lines[1].format)

    def test_external_hyperlink_target_changes_format(self):
        body = p(runs=hyperlink('リンク文字'))
        a = self.first_line(body, rels=[('rLink', 'hyperlink', 'https://example.invalid/a', 'External')])
        b = self.first_line(body, rels=[('rLink', 'hyperlink', 'https://example.invalid/b', 'External')])
        self.assertEqual(a.text, b.text)
        self.assertIn('https://example.invalid/a', a.format)
        self.assertNotEqual(a.format, b.format)

    def test_image_bytes_change_image_token_and_hash(self):
        body = p(runs=image())
        kwargs = {'rels': [('rImage', 'image', 'media/p.png')]}
        a = self.first_line(body, parts={'word/media/p.png': PNG}, **kwargs)
        b = self.first_line(body, parts={'word/media/p.png': PNG + b'changed'}, **kwargs)
        self.assertIn(hashlib.sha256(PNG).hexdigest()[:8], a.text)
        self.assertNotEqual(a.text, b.text)
        self.assertNotEqual(a.format, b.format)

    def test_field_code_deleted_text_and_ruby_are_hashed_verbatim(self):
        variants = (
            ('<w:fldSimple w:instr="REF TargetA">' + run('同じ結果') + '</w:fldSimple>',
             '<w:fldSimple w:instr="REF TargetB">' + run('同じ結果') + '</w:fldSimple>'),
            (tracked('旧甲', 'del') + run('現行'), tracked('旧乙', 'del') + run('現行')),
            ('<w:r><w:ruby><w:rt>' + run('こう') + '</w:rt><w:rubyBase>' + run('項') + '</w:rubyBase></w:ruby></w:r>',
             '<w:r><w:ruby><w:rt>' + run('うなじ') + '</w:rt><w:rubyBase>' + run('項') + '</w:rubyBase></w:ruby></w:r>'),
        )
        for before, after in variants:
            with self.subTest(before=before):
                a, b = self.first_line(p(runs=before)), self.first_line(p(runs=after))
                self.assertEqual(a.text, b.text)
                self.assertNotEqual(a.format, b.format)

    def test_run_spacing_boundary_move_changes_hash(self):
        before = self.first_line(p(runs=run('甲', '<w:spacing w:val="20"/>') + run('乙丙')))
        after = self.first_line(p(runs=run('甲乙', '<w:spacing w:val="20"/>') + run('丙')))
        self.assertEqual(before.text, after.text)
        self.assertNotEqual(before.format, after.format)


class UnitTests(WordTestCase):
    def test_markers_attach_to_the_correct_adjacent_block(self):
        body = '<w:bookmarkStart w:id="1" w:name="_Loose"/>' + p('甲') + '<w:bookmarkEnd w:id="1"/>' + p('乙')
        units = self.document(body).stories[0].units
        self.assertEqual((units[0].start, units[0].stop), (0, 3))
        self.assertEqual((units[1].start, units[1].stop), (3, 4))
        self.assertEqual([u.lines[0].text for u in units[:2]], ['甲', '乙'])

    def test_first_end_and_start_before_final_section_attach_to_paragraph(self):
        body = '<w:bookmarkEnd w:id="7"/>' + p('甲') + '<w:bookmarkStart w:id="7" w:name="_Loose"/>'
        units = self.document(body).stories[0].units
        self.assertEqual((units[0].start, units[0].stop), (0, 3))
        self.assertEqual(units[-1].kind, 'section')

    def test_visible_bookmark_and_permission_ranges_group_paragraphs(self):
        for begin, end in (
            ('<w:bookmarkStart w:id="6" w:name="VisibleTarget"/>', '<w:bookmarkEnd w:id="6"/>'),
            ('<w:permStart w:id="6" w:edGrp="everyone"/>', '<w:permEnd w:id="6"/>'),
        ):
            with self.subTest(begin=begin):
                units = self.document(begin + p('甲') + p('乙') + end + p('丙')).stories[0].units
                self.assertEqual(units[0].kind, 'group')
                self.assertEqual([line.text for line in units[0].lines], ['甲', '乙'])
                self.assertEqual(units[1].lines[0].text, '丙')

    def test_complex_toc_field_groups_across_paragraphs(self):
        body = p(runs='<w:r><w:fldChar w:fldCharType="begin"/><w:instrText>TOC</w:instrText><w:fldChar w:fldCharType="separate"/></w:r>')
        body += p('合成の目次') + p(runs='<w:r><w:fldChar w:fldCharType="end"/></w:r>')
        body += p('独立した段落')
        units = self.document(body).stories[0].units
        self.assertEqual(units[0].kind, 'group')
        self.assertEqual(len(units[0].lines), 3)
        self.assertEqual(units[1].lines[0].text, '独立した段落')

    def test_hidden_bookmark_does_not_group_paragraphs(self):
        units = self.document('<w:bookmarkStart w:id="3" w:name="_HlkSynthetic"/>' + p('甲') + p('乙') + '<w:bookmarkEnd w:id="3"/>').stories[0].units
        self.assertEqual([unit.kind for unit in units], ['paragraph', 'paragraph', 'section'])

    def test_unclosed_complex_field_groups_to_the_end_and_is_unsupported(self):
        body = p(runs='<w:r><w:fldChar w:fldCharType="begin"/><w:instrText>TOC</w:instrText></w:r>') + p('末尾')
        unit = self.document(body).stories[0].units[0]
        self.assertEqual(unit.kind, 'group')
        self.assertIn('閉じていないフィールド', unit.reasons)

    def test_paragraph_section_properties_are_detached(self):
        first = p('区切りの前', props='<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>')
        units = self.document(first + p('次の区間')).stories[0].units
        self.assertEqual([u.kind for u in units], ['paragraph', 'section', 'paragraph', 'section'])
        self.assertEqual(units[1].host, 0)
        self.assertIsNone(units[-1].host)
        self.assertEqual(units[0].lines[0], self.first_line(p('区切りの前')))

    def test_table_shape_signature_and_flattened_lines(self):
        unit = self.document(table([['甲', '乙'], ['丙', '丁']]) + p()).stories[0].units[0]
        self.assertEqual(unit.kind, 'table')
        self.assertIsNotNone(unit.table.signature)
        self.assertEqual([line.kind for line in unit.lines], ['table', 'row', 'cell', 'cell', 'row', 'cell', 'cell'])
        self.assertEqual(unit.lines[0].text, '2×2')
        self.assertEqual([row.line.label for row in unit.table.rows], ['R1', 'R2'])
        self.assertEqual([line.label for line in unit.table.rows[1].cells], ['R2C1', 'R2C2'])
        self.assertEqual(unit.table.rows[0].tc_indices, (0, 1))

    def test_cross_cell_range_and_row_marker_make_table_unsliceable(self):
        cross = table([[p(runs='<w:bookmarkStart w:id="4" w:name="CrossCell"/>' + run('甲')),
                        p(runs=run('乙') + '<w:bookmarkEnd w:id="4"/>')]])
        marker = table([['甲']]).replace('<w:tr>', '<w:bookmarkStart w:id="7" w:name="_Row"/><w:tr>', 1).replace('</w:tr>', '</w:tr><w:bookmarkEnd w:id="7"/>', 1)
        for body in (cross, marker):
            with self.subTest(body=body):
                unit = self.document(body + p()).stories[0].units[0]
                self.assertIsNone(unit.table.signature)

    def test_table_geometry_row_properties_and_cell_format_are_partitioned(self):
        body = table([['甲', '乙']]) + p()
        before = self.document(body).stories[0].units[0]
        variants = (
            (body.replace('<w:tcW w:w="2000"', '<w:tcW w:w="2400"', 1), 0),
            (body.replace('<w:tr>', '<w:tr><w:trPr><w:trHeight w:val="480"/></w:trPr>', 1), 1),
            (body.replace('</w:tcPr>', '<w:shd w:fill="FFEEAA"/></w:tcPr>', 1), 2),
        )
        for changed_body, changed_index in variants:
            with self.subTest(changed_index=changed_index):
                after = self.document(changed_body).stories[0].units[0]
                changed = [i for i, pair in enumerate(zip(before.lines, after.lines)) if pair[0] != pair[1]]
                self.assertEqual(changed, [changed_index])

    def test_table_cell_controls_and_tracked_cell_marks_are_not_sliceable(self):
        control = '<w:sdt><w:sdtPr/><w:sdtContent>' + p('甲') + '</w:sdtContent></w:sdt>'
        cells = table([[control]])
        tracked_cell = table([['甲']], cell_props={(0, 0): '<w:cellIns w:id="8"/>'})
        for body in (cells, tracked_cell):
            with self.subTest(body=body):
                unit = self.document(body + p()).stories[0].units[0]
                self.assertIsNone(unit.table.signature)
        mixed = table([[control + p('末尾')]]) + p()
        changed = mixed.replace('甲', '乙')
        a, b = self.document(mixed).stories[0].units[0], self.document(changed).stories[0].units[0]
        self.assertEqual(a.lines[2].text, '甲\n末尾')
        self.assertEqual(b.lines[2].text, '乙\n末尾')
        self.assertNotEqual(a.lines[2], b.lines[2])
        for member in ('tr', 'tc'):
            wrapper_start = '<w:sdt><w:sdtPr><w:tag w:val="Wrapped' + member + '"/></w:sdtPr><w:sdtContent>'
            raw = table([['包まれたセル']])
            wrapped = raw.replace('<w:' + member + '>', wrapper_start + '<w:' + member + '>', 1)
            wrapped = wrapped.replace('</w:' + member + '>', '</w:' + member + '></w:sdtContent></w:sdt>', 1)
            unit = self.document(wrapped + p()).stories[0].units[0]
            self.assertIsNone(unit.table.signature)
            self.assertEqual(unit.lines[0].text, '1×1')
            self.assertEqual(unit.lines[2].text, '包まれたセル')
            changed = self.document(wrapped.replace('包まれたセル', '書き換えたセル') + p()).stories[0].units[0]
            self.assertNotEqual(unit.key, changed.key)

    def test_group_containing_a_section_is_intrinsically_unsupported(self):
        body = '<w:bookmarkStart w:id="5" w:name="SectionRange"/>'
        body += p('甲', props='<w:sectPr/>') + p('乙') + '<w:bookmarkEnd w:id="5"/>' + p('丙')
        unit = self.document(body).stories[0].units[0]
        self.assertEqual(unit.kind, 'group')
        self.assertIn('セクション区切りを含むまとまり', unit.reasons)

    def test_body_level_revision_is_an_other_unit(self):
        units = self.document(tracked('独立した挿入') + p('末尾')).stories[0].units
        self.assertEqual(units[0].kind, 'other')
        self.assertEqual(units[0].lines[0].label, 'ins')

    def test_role_paired_header_ids_survive_part_renaming(self):
        a = self.document(p('本文'), headers={'default': ('word/header1.xml', p('ヘッダー')), 'first': p('先頭')}, footers={'even': p('フッター')})
        b = self.document(p('本文'), headers={'default': ('word/header8.xml', p('ヘッダー')), 'first': ('word/header9.xml', p('先頭'))}, footers={'even': ('word/footer6.xml', p('フッター'))})
        self.assertEqual([s.id for s in a.stories], [s.id for s in b.stories])
        self.assertEqual([s.title for s in a.stories], ['本文', 'ヘッダー（セクション 1・既定）', 'ヘッダー（セクション 1・先頭ページ）', 'フッター（セクション 1・偶数ページ）'])
        self.assertEqual([s.units[0].lines for s in a.stories], [s.units[0].lines for s in b.stories])
        self.assertEqual(a.stories[1].id, 'header:1:default')

    def test_document_public_parts_and_fresh_trees(self):
        data = make_docx(p('本文'), comment=b'synthetic archive comment')
        doc = self.word.load(data)
        self.assertEqual(doc.data, data)
        self.assertEqual(doc.main, 'word/document.xml')
        self.assertEqual(doc.comment, b'synthetic archive comment')
        self.assertEqual(tuple(info.filename for info in doc.infos), tuple(doc.parts))
        self.assertEqual(doc.content_types.resolve(doc.main), MAIN_TYPES['docx'])
        self.assertEqual(doc.part_sha(doc.main), hashlib.sha256(doc.parts[doc.main]).hexdigest())
        self.assertIsNot(doc.fresh_tree(doc.main).getroot(), doc.fresh_tree(doc.main).getroot())
        rel = doc.rels(doc.main)['rStyles']
        self.assertFalse(rel.external)
        self.assertEqual(rel.path, 'word/styles.xml')

    def test_top_level_visible_bookmark_names_affect_table_and_control_formats(self):
        control = '<w:sdt><w:sdtPr><w:tag w:val="Synthetic"/></w:sdtPr><w:sdtContent>' + p('内容') + '</w:sdtContent></w:sdt>'
        for block in (table([['内容']]), control):
            with self.subTest(block=block):
                first = '<w:bookmarkStart w:id="8" w:name="VisibleA"/>' + block + '<w:bookmarkEnd w:id="8"/>' + p('末尾')
                second = first.replace('VisibleA', 'VisibleB')
                a = self.document(first).stories[0].units[0].lines[0]
                b = self.document(second).stories[0].units[0].lines[0]
                self.assertEqual(a.text, b.text)
                self.assertNotEqual(a.format, b.format)

    def test_field_reference_code_split_between_paragraphs_is_in_unit_refs(self):
        body = p(runs='<w:r><w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve"> REF </w:instrText></w:r>')
        body += p(runs='<w:r><w:instrText>TargetAcrossParagraphs</w:instrText><w:fldChar w:fldCharType="separate"/><w:t>表示結果</w:t><w:fldChar w:fldCharType="end"/></w:r>')
        unit = self.document(body).stories[0].units[0]
        self.assertEqual(unit.kind, 'group')
        self.assertTrue(any(name == 'TargetAcrossParagraphs' for _, name in unit.refs.ref_names))

    def test_unit_refs_keep_fixup_inputs_hidden_by_normalization(self):
        body = '<w:bookmarkStart w:id="30" w:name="_HiddenTarget"/>'
        body += p(style='1', props='<w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr>',
                  runs=run('甲', '<w:rStyle w:val="a0"/>') + image() + hyperlink('外部') + hyperlink('内部', anchor='_HiddenTarget'))
        body += '<w:bookmarkEnd w:id="30"/>'
        doc = self.document(body, numbering=numbering_xml(), rels=[('rImage', 'image', 'media/pixel.png'), ('rLink', 'hyperlink', 'https://example.invalid/test', 'External')], parts={'word/media/pixel.png': PNG})
        refs = doc.stories[0].units[0].refs
        self.assertIn(('paragraph', 'heading 1'), refs.styles)
        self.assertIn(('character', 'Default Paragraph Font'), refs.styles)
        self.assertEqual(refs.num_ids, ('7',))
        self.assertEqual(set(refs.rel_ids), {'rImage', 'rLink'})
        self.assertIn('_HiddenTarget', refs.bookmarks)
        self.assertTrue(any(name == '_HiddenTarget' for _, name in refs.ref_names))
        style = doc.styles['1']
        self.assertEqual((style.id, style.type, style.name, style.based_on), ('1', 'paragraph', 'heading 1', 'a'))
        number = doc.numbering['7']
        self.assertEqual(number.abstract_id, '3')
        self.assertEqual(number.levels[0] if 0 in number.levels else number.levels['0'], ('decimal', '%1.'))
        self.assertRegex(number.def_hash, r'^[0-9a-f]+$')

    def test_bookmarks_and_split_field_reference_names_are_collected(self):
        body = '<w:bookmarkStart w:id="9" w:name="TargetName"/>' + p('参照先') + '<w:bookmarkEnd w:id="9"/>'
        body += p(runs=hyperlink('文書内リンク', anchor='TargetName'))
        body += p(runs='<w:r><w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve"> REF </w:instrText></w:r><w:r><w:instrText>TargetName</w:instrText><w:fldChar w:fldCharType="separate"/><w:t>参照先</w:t><w:fldChar w:fldCharType="end"/></w:r>')
        doc = self.document(body)
        self.assertIn('targetname', doc.bookmarks)
        self.assertGreaterEqual(sum(name.casefold() == 'targetname' for _, _, name in doc.ref_names), 2)



class IntrinsicReasonTests(WordTestCase):
    def test_intrinsic_unsupported_reasons_are_kept_per_unit_and_still_render(self):
        control_begin = '<w:sdt><w:sdtPr>'
        control_end = '</w:sdtPr><w:sdtContent>' + p('制御の合成本文') + '</w:sdtContent></w:sdt>'
        drawing_begin = '<w:r><w:drawing><a:graphic><a:graphicData>'
        drawing_end = '</a:graphicData></a:graphic></w:drawing></w:r>'
        cases = (
            ('コメント', p(runs='<w:r><w:commentReference w:id="4"/></w:r>')),
            ('脚注・文末脚注', p(runs='<w:r><w:footnoteReference w:id="4"/></w:r>')),
            ('脚注・文末脚注', p(runs='<w:r><w:endnoteReference w:id="4"/></w:r>')),
            ('グラフ', p(runs=drawing_begin + '<c:chart r:id="rChart"/>' + drawing_end)),
            ('SmartArt', p(runs=drawing_begin + '<dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rDiagram"/>' + drawing_end)),
            ('図形・テキストボックス', p(runs=drawing_begin + '<wps:wsp><wps:txbx><w:txbxContent>' + p('図形の本文') + '</w:txbxContent></wps:txbx></wps:wsp>' + drawing_end)),
            ('VML の図形・画像', p(runs='<w:r><w:pict><v:shape/></w:pict></w:r>')),
            ('埋め込みオブジェクト（OLE）', p(runs='<w:r><w:object><o:OLEObject/></w:object></w:r>')),
            ('ActiveX コントロール', p(runs='<w:r><w:control/></w:r>')),
            ('インク・3D モデル', p(runs='<w:r><w:contentPart/></w:r>')),
            ('変更履歴の移動', p(runs=tracked('移動元', 'moveFrom') + tracked('移動先', 'moveTo', ident=6))),
            ('データにリンクしたコンテンツ コントロール', control_begin + '<w:dataBinding w:storeItemID="synthetic-store" w:xpath="/test/value"/>' + control_end),
            ('繰り返しセクション', control_begin + '<w15:repeatingSection xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"/>' + control_end),
            ('プレースホルダー表示中のコンテンツ コントロール', control_begin + '<w:showingPlcHdr/><w:placeholder><w:docPart w:val="SyntheticPlaceholder"/></w:placeholder>' + control_end),
            ('埋め込み文書（altChunk）', '<w:altChunk r:id="rChunk"/>'),
            ('サブ文書', p(runs='<w:subDoc r:id="rSubDocument"/>')),
            ('カスタム XML', '<w:customXml w:element="synthetic">' + p('旧形式の合成本文') + '</w:customXml>'),
            ('未対応の要素 <w:futureWidget>', p(runs='<w:r><w:futureWidget/></w:r>')),
        )
        for reason, body in cases:
            with self.subTest(reason=reason):
                data = make_docx(body + p('次の独立した段落'))
                doc = self.word.load(data)
                self.assertIn(reason, doc.stories[0].units[0].reasons)
                self.assertEqual(doc.stories[0].units[-2].reasons, ())
                self.assertTrue(self.word.render_text(data).startswith('# 本文\n'))

    def test_section_inside_a_table_is_intrinsically_unsupported(self):
        cell = p('セル', props='<w:sectPr/>')
        unit = self.document(table([[cell]]) + p()).stories[0].units[0]
        self.assertIn('セクション区切りを含むまとまり', unit.reasons)

    def test_placeholder_reason_depends_on_the_available_glossary_part(self):
        control = '<w:sdt><w:sdtPr><w:showingPlcHdr/><w:placeholder><w:docPart w:val="SyntheticPlaceholder"/></w:placeholder></w:sdtPr><w:sdtContent>' + p('合成プレースホルダー') + '</w:sdtContent></w:sdt>'
        glossary = xml_document('glossaryDocument', '<w:docParts><w:docPart><w:docPartPr><w:name w:val="SyntheticPlaceholder"/></w:docPartPr><w:docPartBody>' + p('登録された合成部品') + '</w:docPartBody></w:docPart></w:docParts>')
        doc = self.document(control + p(), rels=[('rGlossary', 'glossaryDocument', 'glossary/document.xml')], parts={'word/glossary/document.xml': glossary})
        self.assertNotIn('プレースホルダー表示中のコンテンツ コントロール', doc.stories[0].units[0].reasons)
        self.assertEqual([s.id for s in doc.stories], ['body'])
        self.assertIn('# 文書パーツ: SyntheticPlaceholder', self.word.render_text(doc.data))

    def test_unknown_word_elements_in_properties_and_hyperlinks_are_unsupported(self):
        for body in (p(runs=run('本文', '<w:futureProperty/>')),
                     p(runs='<w:hyperlink w:anchor="Target"><w:futureHyperlinkItem/>' + run('本文') + '</w:hyperlink>')):
            with self.subTest(body=body):
                reasons = self.document(body).stories[0].units[0].reasons
                self.assertTrue(any(reason.startswith('未対応の要素 <w:future') for reason in reasons), reasons)

    def test_word_picture_blip_extension_is_supported_and_keeps_image_references(self):
        extension = ('<a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">'
                     '<a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>'
                     '</a:ext></a:extLst>')
        drawing = image().replace('<a:blip r:embed="rImage"/>',
                                  '<a:blip r:embed="rImage">' + extension + '</a:blip>')
        unit = self.document(p(runs=drawing), rels=[('rImage', 'image', 'media/pixel.png')],
                             parts={'word/media/pixel.png': PNG}).stories[0].units[0]
        self.assertEqual(unit.reasons, ())
        self.assertEqual(unit.refs.rels, ('rImage',))
        self.assertIn('［画像 ', unit.lines[0].text)

    def test_drawing_extensions_keep_effects_and_all_relationship_references(self):
        def picture(effect):
            extension = ('<a:extLst><a:ext uri="synthetic-picture-effects">'
                         '<a14:imgProps xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">'
                         '<a14:imgLayer r:embed="rEffect"><a14:brightnessContrast bright="'
                         + effect + '"/></a14:imgLayer></a14:imgProps></a:ext></a:extLst>')
            return image().replace('<a:blip r:embed="rImage"/>',
                                   '<a:blip r:embed="rImage">' + extension + '</a:blip>')
        kwargs = {'rels': [('rImage', 'image', 'media/pixel.png'),
                           ('rEffect', 'hdphoto', 'media/effect.wdp')],
                  'parts': {'word/media/pixel.png': PNG, 'word/media/effect.wdp': b'synthetic-effect'}}
        first = self.document(p(runs=picture('0')), **kwargs).stories[0].units[0]
        changed = self.document(p(runs=picture('2000')), **kwargs).stories[0].units[0]
        self.assertEqual(first.reasons, ())
        self.assertEqual(first.refs.rels, ('rImage', 'rEffect'))
        self.assertNotEqual(first.lines[0].format, changed.lines[0].format)

    def test_drawing_extension_container_does_not_allow_object_payloads(self):
        payloads = (('グラフ', '<c:chart r:id="rChart"/>'),
                    ('図形・テキストボックス', '<wps:wsp/>'),
                    ('VML の図形・画像', '<v:shape/>'),
                    ('インク・3D モデル', '<w:contentPart/>'))
        for reason, payload in payloads:
            with self.subTest(reason=reason):
                drawing = image().replace('<a:blip r:embed="rImage"/>',
                    '<a:blip r:embed="rImage"><a:extLst><a:ext uri="synthetic-payload">'
                    + payload + '</a:ext></a:extLst></a:blip>')
                self.assertIn(reason, self.document(p(runs=drawing)).stories[0].units[0].reasons)

    def test_microsoft_word_and_relative_picture_properties_are_supported(self):
        controls = ('<w:sdt><w:sdtPr><w15:appearance xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" w15:val="tags"/>'
                    '<w15:color xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" w15:val="808080"/>'
                    '</w:sdtPr><w:sdtContent>' + p('制御本文') + '</w:sdtContent></w:sdt>')
        paragraphs = p('段落設定', props='<w15:collapsed xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"/>')
        paragraphs += p(runs=run('文字設定', '<w14:glow w14:rad="12000"><w14:srgbClr w14:val="FF0000"/></w14:glow>'
                                 '<w16:property xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" w16:val="1"/>'))
        drawing = image().replace('</wp:inline>',
            '<wp14:sizeRelH relativeFrom="page"><wp14:pctWidth>50000</wp14:pctWidth></wp14:sizeRelH>'
            '<wp14:sizeRelV relativeFrom="page"><wp14:pctHeight>50000</wp14:pctHeight></wp14:sizeRelV></wp:inline>')
        doc = self.document(controls + paragraphs + p(runs=drawing))
        self.assertEqual([unit.reasons for unit in doc.stories[0].units], [()] * len(doc.stories[0].units))

    def test_images_links_and_plain_tracked_insertions_or_deletions_are_supported(self):
        body = p(runs=image() + hyperlink() + tracked('追加') + tracked('削除', 'del'))
        unit = self.document(body, rels=[('rImage', 'image', 'media/pixel.png'), ('rLink', 'hyperlink', 'https://example.invalid/test', 'External')], parts={'word/media/pixel.png': PNG}).stories[0].units[0]
        self.assertEqual(unit.reasons, ())


class LoadRefusalTests(WordTestCase):
    def assert_refusal(self, data, text, purpose='edit'):
        with self.assertRaisesRegex(self.word.DocumentError, re.escape(text)):
            self.word.load(data, purpose=purpose)

    def test_ole_is_refused_before_zip_parsing(self):
        self.assert_refusal(bytes.fromhex('D0CF11E0A1B11AE1') + b'not a zip', '暗号化された Word 文書（パスワード付き）または .doc 形式は扱えません')

    def test_bad_zip_and_duplicate_names(self):
        self.assert_refusal(b'PK broken archive', 'Word 文書を読めません:')
        output = io.BytesIO(make_docx(p('本文')))
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', UserWarning)
            with zipfile.ZipFile(output, 'a') as archive:
                archive.writestr('word/document.xml', b'<duplicate/>')
        self.assert_refusal(output.getvalue(), '同名パーツが重複する Word 文書は扱えません')

    def test_dtd_is_refused(self):
        data = make_docx(p('本文'))
        parts = unpack(data)
        parts['word/document.xml'] = parts['word/document.xml'].replace(b'?>', b'?><!DOCTYPE document [<!ENTITY text "blocked">]>', 1)
        self.assert_refusal(zip_parts(parts), 'DTD を含む XML は扱えません: word/document.xml')

    def test_missing_main_part_and_ambiguous_or_external_main_relationship(self):
        parts = unpack(make_docx(p('本文')))
        del parts['word/document.xml']
        self.assert_refusal(zip_parts(parts), 'Word')
        for rels in ([], [('one', 'officeDocument', 'word/document.xml'), ('two', 'officeDocument', 'word/document.xml')],
                     [('one', 'officeDocument', 'https://example.invalid/document.xml', 'External')]):
            with self.subTest(rels=rels):
                self.assert_refusal(make_docx(p('本文'), root_rels=rels), 'Word の本文パーツの参照が不正です')

    def test_strict_namespace_relationship_and_conformance(self):
        original = make_docx(p('本文'))
        for part, old, new in (
            ('word/document.xml', W, STRICT_W),
            ('_rels/.rels', R + '/officeDocument', STRICT_R + '/officeDocument'),
            ('word/document.xml', '<w:document ', '<w:document w:conformance="strict" '),
        ):
            with self.subTest(part=part, old=old):
                parts = unpack(original)
                parts[part] = parts[part].replace(old.encode(), new.encode())
                self.assert_refusal(zip_parts(parts), 'Strict 形式の Word 文書は扱えません')

    def test_non_word_main_content_type_is_refused(self):
        self.assert_refusal(make_docx(p('本文'), main_type='application/xml'), 'Word 文書ではありません（本文パーツの種類: application/xml）')

    def test_signatures_are_edit_only_refusals(self):
        data = make_docx(p('署名付きの合成例'), signature=True)
        self.assert_refusal(data, '電子署名付きの Word 文書は段落単位で反映できません')
        self.assertIn('署名付きの合成例', self.word.render_text(data))
        origin = 'http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin'
        signed = make_docx(p('本文'), root_rels=[('main', 'officeDocument', 'word/document.xml'), ('signature', origin, '_xmlsignatures/origin.sigs')])
        self.assert_refusal(signed, '電子署名付きの Word 文書は段落単位で反映できません')

    def test_refusal_order_keeps_xml_and_main_type_checks_before_signatures(self):
        parts = unpack(make_docx(p('本文'), signature=True))
        parts['word/document.xml'] = parts['word/document.xml'].replace(b'?>', b'?><!DOCTYPE document>', 1)
        self.assert_refusal(zip_parts(parts), 'DTD を含む XML は扱えません: word/document.xml')
        self.assert_refusal(make_docx(p('本文'), main_type='application/xml', signature=True),
                            'Word 文書ではありません（本文パーツの種類: application/xml）')
        duplicate = [('same', 'hyperlink', 'https://example.invalid/a', 'External'),
                     ('same', 'hyperlink', 'https://example.invalid/b', 'External')]
        self.assert_refusal(make_docx(p('本文'), signature=True, rels=duplicate),
                            '電子署名付きの Word 文書は段落単位で反映できません')

    def test_duplicate_relationship_identifiers(self):
        data = make_docx(p('本文'), rels=[('duplicate', 'hyperlink', 'https://example.invalid/a', 'External'), ('duplicate', 'hyperlink', 'https://example.invalid/b', 'External')])
        self.assert_refusal(data, 'relationship ID が重複しています: word/document.xml')

    def test_duplicate_bookmark_ids_or_names_are_edit_only_refusals(self):
        for second_id, second_name in ((1, 'TargetB'), (2, 'TARGETA')):
            with self.subTest(second_id=second_id, second_name=second_name):
                body = '<w:bookmarkStart w:id="1" w:name="TargetA"/>' + p('甲') + '<w:bookmarkEnd w:id="1"/>'
                body += '<w:bookmarkStart w:id="' + str(second_id) + '" w:name="' + second_name + '"/>' + p('乙') + '<w:bookmarkEnd w:id="' + str(second_id) + '"/>'
                data = make_docx(body)
                self.assert_refusal(data, 'ブックマークの ID が重複しています: 本文')
                self.assertIn('甲', self.word.render_text(data))

    def test_all_four_word_content_types_and_macro_part_are_accepted(self):
        self.assertEqual(self.word.EXTENSIONS, {'.docx', '.docm', '.dotx', '.dotm'})
        self.assertEqual(self.word.MAIN_TYPES, set(MAIN_TYPES.values()))
        for kind in MAIN_TYPES:
            with self.subTest(kind=kind):
                doc = self.document(p('合成文書'), main_type=kind)
                self.assertEqual(doc.content_types.resolve(doc.main), MAIN_TYPES[kind])
                if kind in ('docm', 'dotm'):
                    self.assertIn('word/vbaProject.bin', doc.parts)


class ListMarkerSummaryTests(WordTestCase):
    def list_data(self, text, fonts='', fmt='bullet'):
        numbering = numbering_xml(bullet=True).replace('・', text).replace('w:val="bullet"', 'w:val="' + fmt + '"')
        if fonts:
            numbering = numbering.replace('</w:lvl>', '<w:rPr><w:rFonts ' + fonts + '/></w:rPr></w:lvl>')
        body = p('一覧の項目', props='<w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr>')
        return make_docx(body, numbering=numbering)

    def test_symbol_font_marker_mappings_are_readable_and_deterministic(self):
        cases = (
            ('Symbol', 0xB7, '•'),
            ('Wingdings', 0x6C, '●'), ('Wingdings', 0x6E, '■'), ('Wingdings', 0x6F, '□'),
            ('Wingdings', 0x71, '❑'), ('Wingdings', 0x75, '◆'), ('Wingdings', 0x76, '❖'),
            ('Wingdings', 0xA7, '▪'), ('Wingdings', 0xA8, '◻'), ('Wingdings', 0xD8, '➢'),
            ('Wingdings', 0xFC, '✔'),
        )
        for font, code, expected in cases:
            with self.subTest(font=font, code=code):
                marker = chr(0xF000 + code)
                data = self.list_data(marker, 'w:ascii="' + font + '"')
                doc = self.word.load(data)
                line = doc.stories[0].units[0].lines[0]
                self.assertIn('箇条書き「' + expected + '」L1', line.format)
                self.assertEqual(doc.numbering['7'].levels['0'], ('bullet', marker))
                text = self.word.render_text(data)
                self.assertIn('箇条書き「' + expected + '」L1', text)
                self.assertEqual(text, self.word.render_text(data))

    def test_unknown_marker_or_font_uses_lowercase_unicode_escape(self):
        cases = (
            ('\uf0e0', 'w:ascii="Wingdings"', r'\uf0e0'),
            ('\uf06c', '', r'\uf06c'),
            ('\uf06c', 'w:ascii="Unknown"', r'\uf06c'),
        )
        for marker, fonts, expected in cases:
            with self.subTest(marker=marker, fonts=fonts):
                data = self.list_data(marker, fonts)
                line = self.lines(data)[0]
                self.assertIn('箇条書き「' + expected + '」L1', line.format)
                self.assertEqual(len(expected), 6)
                self.assertNotIn(marker, line.format)
                rendered = self.word.render_text(data)
                self.assertIn('箇条書き「' + expected + '」L1', rendered)
                self.assertEqual(re.search('箇条書き「([^」]*)」', rendered).group(1), expected)

    def test_marker_font_is_case_insensitive_and_ascii_precedes_hansi(self):
        cases = (
            ('\uf06c', 'w:ascii="wInGdInGs"', '●'),
            ('\uf0b7', 'w:hAnsi="sYmBoL"', '•'),
            ('\uf06c', 'w:ascii="Unknown" w:hAnsi="Wingdings"', r'\uf06c'),
            ('\uf0b7', 'w:ascii="Symbol" w:hAnsi="Wingdings"', '•'),
        )
        for marker, fonts, expected in cases:
            with self.subTest(fonts=fonts):
                line = self.lines(self.list_data(marker, fonts))[0]
                self.assertIn('箇条書き「' + expected + '」L1', line.format)

    def test_bullet_and_number_summaries_only_replace_symbol_pua_range(self):
        marker = '\uefff\uf000\uf06c\uf0ff\uf100%1.・」L1\uf0e0'
        expected = '\uefff' + r'\uf000' + '●' + r'\uf0ff' + '\uf100%1.・」L1' + r'\uf0e0'
        for fmt, label in (('bullet', '箇条書き'), ('decimal', '番号')):
            with self.subTest(fmt=fmt):
                data = self.list_data(marker, 'w:ascii="Wingdings"', fmt)
                line = self.lines(data)[0]
                self.assertIn(label + '「' + expected + '」L1', line.format)
                self.assertIn(label + '「' + expected + '」L1', self.word.render_text(data))

    def test_marker_display_preserves_other_summary_and_label_escaping(self):
        text = r'箇条書き「\uf0e0」L1' + '\n' + r'C:\folder'
        self.assertEqual(self.word._safe_format(text), json.dumps(text, ensure_ascii=False)[1:-1])
        line = self.word.Line('paragraph', text, text, r'番号「\uf0e0」L1' + '\n' + r'C:\folder')
        rendered = self.word._text_line(line)
        self.assertTrue(rendered.startswith('[' + self.word._safe_format(text) + '] '))
        self.assertIn('"' + self.word._safe_format(text) + '"', rendered)
        self.assertTrue(rendered.endswith(r'番号「\uf0e0」L1\nC:\\folder'))
        self.assertEqual(len(rendered.splitlines()), 1)

    def test_each_level_and_level_override_use_their_own_marker_font(self):
        numbering = numbering_xml(bullet=True)
        numbering = numbering.replace('<w:lvlText w:val="・"/>', '<w:lvlText w:val="\uf06c"/><w:rPr><w:rFonts w:ascii="Wingdings"/></w:rPr>', 1)
        numbering = numbering.replace('<w:lvlText w:val="・"/>', '<w:lvlText w:val="\uf0b7"/><w:rPr><w:rFonts w:ascii="Symbol"/></w:rPr>', 1)
        overrides = (
            ('0', '', '●'), ('1', '', '•'),
            ('0', '<w:startOverride w:val="2"/>', '●'),
            ('0', '<w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="\uf06e"/><w:rPr><w:rFonts w:hAnsi="Wingdings"/></w:rPr></w:lvl>', '■'),
            ('0', '<w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="\uf06e"/></w:lvl>', r'\uf06e'),
        )
        for level, override, expected in overrides:
            with self.subTest(level=level, override=override):
                changed = numbering
                if override:
                    changed = changed.replace('</w:num>', '<w:lvlOverride w:ilvl="0">' + override + '</w:lvlOverride></w:num>')
                body = p('一覧', props='<w:numPr><w:ilvl w:val="' + level + '"/><w:numId w:val="7"/></w:numPr>')
                line = self.first_line(body, numbering=changed)
                self.assertIn('箇条書き「' + expected + '」L' + str(int(level) + 1), line.format)

    def test_readable_marker_does_not_normalize_comparison_data(self):
        pua = self.word.load(self.list_data('\uf06c', 'w:ascii="Wingdings"'))
        unicode = self.word.load(self.list_data('●', 'w:ascii="Wingdings"'))
        a, b = pua.stories[0].units[0], unicode.stories[0].units[0]
        self.assertIn('箇条書き「●」L1', a.lines[0].format)
        self.assertIn('箇条書き「●」L1', b.lines[0].format)
        self.assertEqual(pua.numbering['7'].levels['0'], ('bullet', '\uf06c'))
        self.assertNotEqual(pua.numbering['7'].def_hash, unicode.numbering['7'].def_hash)
        self.assertNotEqual(pua.canon(a.nodes[0], pua.main), unicode.canon(b.nodes[0], unicode.main))
        self.assertNotEqual(pua.fmt_hash(a.nodes[0], pua.main), unicode.fmt_hash(b.nodes[0], unicode.main))


class RecordAndTextconvTests(WordTestCase):
    def test_record_escaping_preserves_one_line_and_five_strings(self):
        fields = ['12', 'paragraph', '見出し', '引用"\\\t\n\u0085\u2028\u2029', '']
        encoded = self.word.record_json(fields)
        self.assertEqual(json.loads(encoded), fields)
        self.assertEqual(len(encoded.splitlines()), 1)
        for raw in ('\t', '\n', '\u0085', '\u2028', '\u2029'):
            self.assertNotIn(raw, encoded)
        for escaped in ('\\t', '\\n', '\\u0085', '\\u2028', '\\u2029'):
            self.assertIn(escaped, encoded)
        line = self.first_line(p('文字'))
        self.assertEqual(json.loads(line.key), [line.kind, line.label, line.text, line.format])

    def test_plain_textconv_output_uses_json_escaping_and_lf(self):
        data = make_docx(p('合成の見出し', style='1') + p(runs=run('本文') + '<w:r><w:tab/><w:t>改行</w:t><w:br/><w:t>続き</w:t></w:r>'))
        text = self.word.render_text(data)
        self.assertTrue(text.startswith('# 本文\n[heading 1] "合成の見出し"\n'))
        self.assertIn('"本文\\t改行\\n続き"', text)
        self.assertNotIn('\r', text)
        self.assertTrue(text.endswith('\n'))
        self.assertEqual(self.word.placeholder_lines(data), ['# 本文'])

    def test_table_textconv_labels_and_nested_table_display(self):
        nested = table([['内甲', '内乙']])
        body = table([[p('外側') + nested + p()]]) + p('末尾')
        text = self.word.render_text(make_docx(body))
        self.assertIn('表 "1×1"', text)
        self.assertIn('  行 R1 ""', text)
        self.assertIn('  R1C1 ', text)
        self.assertIn('［表: 内甲 | 内乙］', text)

    def test_comprehensive_japanese_document_renders_every_supported_display(self):
        notes = '<w:footnote w:id="-1" w:type="separator">' + p('区切りは表示しない') + '</w:footnote><w:footnote w:id="15">' + p('脚注の本文') + '</w:footnote>'
        endnotes = '<w:endnote w:id="20">' + p('文末脚注の本文') + '</w:endnote>'
        comments = '<w:comment w:id="6" w:author="合成者" w:date="2026-03-04T05:06:07Z">' + p('コメントの本文') + '</w:comment>'
        body = p('合成見出し', style='1')
        body += p('箇条書き項目', props='<w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr>')
        body += table([['項目', '数値'], ['合成品', '42']])
        body += p(runs=image() + hyperlink('外部リンク'))
        body += p(runs='<w:r><w:footnoteReference w:id="15"/><w:endnoteReference w:id="20"/><w:commentReference w:id="6"/></w:r>')
        body += p(runs=tracked('削除された試験文字', 'del') + tracked('追加された試験文字'))
        body += p(runs='<w:r><w:ruby><w:rt>' + run('しけん') + '</w:rt><w:rubyBase>' + run('試験') + '</w:rubyBase></w:ruby></w:r>')
        body += p(runs='<mc:AlternateContent><mc:Choice Requires="w16se"><w:r><w16se:symEx w16se:char="1F642" w16se:font="Segoe UI Emoji"/></w:r></mc:Choice><mc:Fallback>' + run('🙂') + '</mc:Fallback></mc:AlternateContent>')
        box = '<w:r><w:drawing><wp:inline><a:graphic><a:graphicData><wps:wsp><wps:txbx><w:txbxContent>' + p('選択側テキストボックス') + '</w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
        body += p(runs='<mc:AlternateContent><mc:Choice Requires="wps">' + box + '</mc:Choice><mc:Fallback><w:r><w:pict><v:shape><v:textbox><w:txbxContent>' + p('重複する代替表示') + '</w:txbxContent></v:textbox></v:shape></w:pict></w:r></mc:Fallback></mc:AlternateContent>')
        data = make_docx(body, numbering=numbering_xml(bullet=True), headers={'default': p('合成ヘッダー')}, footers={'default': p('合成フッター')}, footnotes=notes, endnotes=endnotes, comments=comments,
                         rels=[('rImage', 'image', 'media/pixel.png'), ('rLink', 'hyperlink', 'https://example.invalid/word', 'External')], parts={'word/media/pixel.png': PNG})
        text = self.word.render_text(data)
        for expected in ('# 本文', '[heading 1] "合成見出し"', '箇条書き「・」L1', '［画像 ', 'https://example.invalid/word',
                         '［脚注 1］', '［文末脚注 1］', '［コメント 1］', '# 脚注', '脚注 1 "脚注の本文"', '# 文末脚注', '# コメント', '合成者',
                         '変更履歴(削除', '変更履歴(挿入', 'ルビ「試験(しけん)」', '🙂', 'テキストボックス "選択側テキストボックス"',
                         '# ヘッダー（セクション 1・既定）', '# フッター（セクション 1・既定）'):
            with self.subTest(expected=expected):
                self.assertIn(expected, text)
        self.assertNotIn('重複する代替表示', text)
        self.assertNotIn('区切りは表示しない', text)
        self.assertEqual(text.count('テキストボックス "選択側テキストボックス"'), 1)
        self.assertEqual(text, self.word.render_text(data))

    def test_content_control_header_and_inner_lines(self):
        control = '<w:sdt><w:sdtPr><w:tag w:val="SyntheticTag"/><w:alias w:val="合成別名"/><w:id w:val="123"/></w:sdtPr><w:sdtContent>' + p('制御内の本文') + '</w:sdtContent></w:sdt>'
        text = self.word.render_text(make_docx(control + p('末尾')))
        self.assertIn('コンテンツ コントロール [SyntheticTag] ""', text)
        self.assertIn('  "制御内の本文"', text)

    def test_specification_like_fixture_has_toc_headers_notes_lists_and_visible_edits(self):
        before = self.word.load(specification_docx())
        after = self.word.load(specification_docx(updated=True))
        toc = [unit for unit in before.stories[0].units if unit.kind == 'group']
        self.assertEqual(len(toc), 1)
        self.assertGreater(len(toc[0].lines), 1)
        self.assertGreater(len(before.stories), 2)
        self.assertIn(('footnote', '8'), before.notes)
        self.assertTrue(any('番号「%1.」L1' in line.format for unit in before.stories[0].units for line in unit.lines))
        a = {line.text: line for unit in before.stories[0].units for line in unit.lines if line.kind == 'paragraph'}
        b = {line.text: line for unit in after.stories[0].units for line in unit.lines if line.kind == 'paragraph'}
        self.assertEqual(a['装置の温度を記録する。'], b['装置の温度を記録する。'])
        self.assertNotEqual(a['注意事項'].format, b['注意事項'].format)
        self.assertNotEqual(self.word.render_text(specification_docx()), self.word.render_text(specification_docx(updated=True)))

    def test_field_state_across_paragraphs_hides_pre_result_text_but_hashes_it(self):
        body = p(runs='<w:r><w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve"> REF SyntheticTarget </w:instrText><w:t>非表示甲</w:t></w:r>')
        body += p(runs='<w:r><w:t>非表示乙</w:t><w:fldChar w:fldCharType="separate"/><w:t>表示結果</w:t><w:fldChar w:fldCharType="end"/></w:r>')
        before = self.document(body)
        after = self.document(body.replace('非表示甲', '非表示丙').replace('非表示乙', '非表示丁'))
        first = before.stories[0].units[0]
        second = after.stories[0].units[0]
        self.assertEqual(first.kind, 'group')
        self.assertEqual([line.text for line in first.lines], ['', '表示結果'])
        self.assertEqual([line.text for line in first.lines], [line.text for line in second.lines])
        for a, b in zip(first.lines, second.lines):
            self.assertNotEqual(a.format, b.format)
        text = self.word.render_text(before.data)
        self.assertNotIn('非表示甲', text)
        self.assertNotIn('非表示乙', text)
        self.assertIn('"表示結果"', text)

    def test_format_hash_uses_raw_canonical_json_before_record_escaping(self):
        body = p(runs='<w:r><w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve"> REF Target\u2028Other </w:instrText><w:fldChar w:fldCharType="separate"/><w:t>結果</w:t><w:fldChar w:fldCharType="end"/></w:r>')
        doc = self.document(body)
        node = doc.fresh_tree(doc.main).getroot().find('.//{' + W + '}p')
        canonical = doc.canon(node, doc.main)
        raw = json.dumps(canonical, ensure_ascii=False, separators=(',', ':'))
        expected = hashlib.sha256(raw.encode('utf-8')).hexdigest()[:8]
        self.assertIn('\u2028', raw)
        self.assertNotEqual(raw, self.word.record_json(canonical))
        self.assertEqual(doc.fmt_hash(node, doc.main), expected)
        self.assertTrue(doc.stories[0].units[0].lines[0].format.endswith('#' + expected))

    def test_render_text_is_independent_of_python_hash_seed(self):
        code = ('import sys;sys.path[:0]=[sys.argv[1],sys.argv[2]];'
                'import word_ooxml;from word_fixtures import specification_docx;'
                'sys.stdout.buffer.write(word_ooxml.render_text(specification_docx(updated=True)).encode("utf-8"))')
        results = []
        for seed in ('1', '17', '123456'):
            env = dict(os.environ, PYTHONHASHSEED=seed, PYTHONUTF8='1', PYTHONDONTWRITEBYTECODE='1')
            result = subprocess.run([sys.executable, '-B', '-c', code, str(ROOT / 'scripts'), str(ROOT / 'tests')],
                                    capture_output=True, env=env, timeout=60)
            self.assertEqual(result.returncode, 0, result.stderr.decode('utf-8', errors='replace'))
            results.append(result.stdout)
        self.assertEqual(results[0], results[1])
        self.assertEqual(results[0], results[2])


class ReadModelEdgeTests(WordTestCase):
    def test_hidden_markers_attached_outside_table_preserve_slicing(self):
        plain = table([['甲']]) + p()
        marked = '<w:bookmarkStart w:id="40" w:name="_HiddenTarget"/>' + table([['甲']]) + '<w:bookmarkEnd w:id="40"/>' + p()
        before, after = self.document(plain).stories[0].units[0], self.document(marked).stories[0].units[0]
        self.assertEqual(before.lines, after.lines)
        self.assertEqual(before.table.signature, after.table.signature)
        self.assertIsNotNone(after.table.signature)
        self.assertIn('_HiddenTarget', after.refs.bookmarks)

    def test_content_control_ids_are_noise(self):
        control = '<w:sdt><w:sdtPr><w:tag w:val="SampleControl"/><w:id w:val="77"/></w:sdtPr><w:sdtContent>' + p('同じ本文') + '</w:sdtContent></w:sdt>'
        self.assert_noise(make_docx(control), make_docx(control.replace('w:val="77"', 'w:val="88"')))

    def test_notes_expose_whole_document_references_without_picker_stories(self):
        note = '<w:footnote w:id="8"><w:bookmarkStart w:id="60" w:name="NoteTarget"/>'
        note += p(runs='<w:fldSimple w:instr="REF BodyTarget">' + run('参照結果') + '</w:fldSimple>')
        note += '<w:bookmarkEnd w:id="60"/></w:footnote>'
        body = '<w:bookmarkStart w:id="61" w:name="BodyTarget"/>' + p('本文') + '<w:bookmarkEnd w:id="61"/>'
        doc = self.document(body, footnotes=note)
        self.assertEqual([s.id for s in doc.stories], ['body'])
        self.assertNotIn('notetarget', doc.bookmarks)
        self.assertEqual(doc.all_bookmarks['notetarget'], [('footnotes', '60')])
        self.assertIn(('footnotes', '相互参照', 'BodyTarget'), doc.all_ref_names)
        self.assertEqual([root[0] for root in doc.reference_roots], ['body', 'footnotes'])
        with self.assertRaisesRegex(self.word.DocumentError, 'ブックマークの ID が重複しています: 脚注'):
            self.document(body, footnotes=note.replace('NoteTarget', 'BodyTarget'))

    def test_external_office_relationship_does_not_count_as_a_main_part(self):
        doc = self.document(p('本文'), root_rels=[
            ('rMain', 'officeDocument', 'word/document.xml'),
            ('rExternal', 'officeDocument', 'https://example.invalid/template', 'External'),
        ])
        self.assertEqual(doc.main, 'word/document.xml')

    def test_strict_non_main_relationship_is_refused(self):
        data = make_docx(p('本文'), rels=[('rImage', STRICT_R + '/image', 'media/p.png')], parts={'word/media/p.png': PNG})
        with self.assertRaisesRegex(self.word.DocumentError, 'Strict 形式の Word 文書は扱えません'):
            self.word.load(data)

    def test_numbering_override_picture_bullet_is_reported(self):
        numbering = numbering_xml().replace('</w:num>', '<w:lvlOverride w:ilvl="0"><w:lvl w:ilvl="0"><w:lvlPicBulletId w:val="5"/><w:numFmt w:val="bullet"/><w:lvlText w:val="x"/></w:lvl></w:lvlOverride></w:num>')
        doc = self.document(p('一覧'), numbering=numbering)
        self.assertTrue(doc.numbering['7'].pic_bullet)

    def test_detached_section_references_belong_only_to_the_section(self):
        body = p('前半', props='<w:sectPr/>') + p('後半')
        doc = self.document(body, headers={('default', 1): p('ヘッダー')})
        paragraph, section = doc.stories[0].units[:2]
        self.assertEqual(paragraph.kind, 'paragraph')
        self.assertEqual(paragraph.refs.rels, ())
        self.assertEqual(section.kind, 'section')
        self.assertTrue(section.refs.rels)
        note_positions = '<w:footnotePr><w:pos w:val="pageBottom"/></w:footnotePr><w:endnotePr><w:pos w:val="docEnd"/></w:endnotePr>'
        positioned = self.document(p('本文') + '<w:sectPr>' + note_positions + '</w:sectPr>')
        self.assertEqual(positioned.stories[0].units[-1].reasons, ())

    def test_glossary_uses_its_own_style_definitions(self):
        glossary = xml_document('glossaryDocument', '<w:docParts><w:docPart><w:docPartPr><w:name w:val="SamplePart"/></w:docPartPr><w:docPartBody>' + p('部品本文', style='1') + '</w:docPartBody></w:docPart></w:docParts>')
        styles = default_styles().replace('w:val="heading 1"', 'w:val="部品の独立スタイル"')
        data = make_docx(p('本文'), rels=[('rGlossary', 'glossaryDocument', 'glossary/document.xml')], parts={
            'word/glossary/document.xml': glossary,
            'word/glossary/styles.xml': styles,
            'word/glossary/_rels/document.xml.rels': relationships([('rStyles', 'styles', 'styles.xml')]),
        })
        text = self.word.render_text(data)
        self.assertIn('# 文書パーツ: SamplePart\n[部品の独立スタイル] "部品本文"', text)

    def test_foreign_unknown_element_is_intrinsically_unsupported(self):
        body = p(runs='<w:r><x:unsupported xmlns:x="urn:synthetic:unknown"/></w:r>')
        self.assertIn('未対応の要素 <x:unsupported>', self.document(body).stories[0].units[0].reasons)


if __name__ == '__main__':
    unittest.main()
