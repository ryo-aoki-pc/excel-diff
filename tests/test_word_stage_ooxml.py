"""Word slot projections, partial rebuilds, fix-ups, and validation."""

import copy
import importlib.util
import io
import json
import os
from pathlib import Path
import subprocess
import sys
import time
import unittest
from unittest import mock
import zipfile

from lxml import etree

from word_fixtures import (
    CT, NAMESPACES, PNG, PR, R, W, default_styles, hyperlink, image,
    make_docx, numbering_xml, p, relationships, run, table, tracked,
    unpack, xml_document,
)


ROOT = Path(__file__).resolve().parents[1]
MODULE = ROOT / 'scripts' / 'word_stage_ooxml.py'
NS = dict(NAMESPACES)
NS['pr'] = PR
NS['ct'] = CT
VALIDATION = '作り直した Word 文書の検証に失敗しました。書き込みは行っていません'


def tag(name, namespace=W):
    return '{' + namespace + '}' + name


def numbered(text, ident=7):
    return p(text, props='<w:numPr><w:ilvl w:val="0"/><w:numId w:val="' + str(ident) + '"/></w:numPr>')


def bookmark(text, name, ident=1):
    return ('<w:bookmarkStart w:id="' + str(ident) + '" w:name="' + name + '"/>'
            + p(text) + '<w:bookmarkEnd w:id="' + str(ident) + '"/>')


def content_control(texts, ident=1):
    return ('<w:sdt><w:sdtPr><w:id w:val="' + str(ident) + '"/><w:tag w:val="Synthetic"/>'
            '</w:sdtPr><w:sdtContent>' + ''.join(p(text) for text in texts)
            + '</w:sdtContent></w:sdt>')


def change_parts(data, changes=None, remove=()):
    """Mutate serialized results while retaining the archive metadata."""
    changes = changes or {}
    result = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as source:
        with zipfile.ZipFile(result, 'w') as target:
            for info in source.infolist():
                if info.filename not in remove:
                    target.writestr(copy.copy(info), changes.get(info.filename, source.read(info.filename)))
            target.comment = source.comment
    return result.getvalue()


def xml_bytes(node):
    return etree.tostring(node, encoding='UTF-8', xml_declaration=True, standalone=True)


class WordStageTestCase(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.module_name = '_word_stage_ooxml_tests'
        sys.path.insert(0, str(ROOT / 'scripts'))
        spec = importlib.util.spec_from_file_location(cls.module_name, MODULE)
        cls.stage = importlib.util.module_from_spec(spec)
        sys.modules[cls.module_name] = cls.stage
        spec.loader.exec_module(cls.stage)
        import word_ooxml
        cls.word = word_ooxml

    @classmethod
    def tearDownClass(cls):
        sys.modules.pop(cls.module_name, None)
        sys.path.remove(str(ROOT / 'scripts'))

    def compare(self, before, after, **options):
        return self.stage.compare(make_docx(before, **options), make_docx(after, **options))

    def story(self, comparison, ident='body'):
        return next(story for story in comparison.stories if story.id == ident)

    def records(self, comparison, side='raw_changed', ident='body'):
        return [self.stage.parse_record(line) for line in getattr(self.story(comparison, ident), side).splitlines()]

    def select(self, comparison, predicate, ident='body'):
        story = self.story(comparison, ident)
        lines = [changed if predicate(self.stage.parse_record(changed)) else baseline
                 for baseline, changed in zip(story.baseline.splitlines(), story.raw_changed.splitlines())]
        return comparison.choose(ident, lines, projection='raw_changed')

    def all_selected(self, comparison):
        selected = frozenset()
        for story in comparison.stories:
            selected |= comparison.choose(story.id, story.raw_changed.splitlines(), projection='raw_changed')
        return selected

    def rebuild_all(self, before, after):
        comparison = self.stage.compare(before, after)
        return comparison.rebuild(self.all_selected(comparison))

    def tree(self, data, part='word/document.xml'):
        return etree.fromstring(unpack(data)[part])

    def assert_blocked(self, before, after, reason):
        comparison = self.stage.compare(before, after)
        matches = [item for item in comparison.unsupported if reason in item.reason]
        self.assertTrue(matches, (reason, comparison.unsupported))
        self.assertGreater(comparison.differences, 0)
        for item in matches:
            story = self.story(comparison, item.story_id)
            positions = [index for index, line in enumerate(story.baseline.splitlines())
                         if self.stage.parse_record(line).id == item.slot
                         or self.stage.parse_record(line).id.startswith(item.slot + '/')]
            self.assertTrue(positions)
            for index in positions:
                self.assertEqual(story.baseline.splitlines()[index], story.changed.splitlines()[index])
        with self.assertRaises(self.word.DocumentError) as error:
            comparison.rebuild(frozenset((item.story_id, item.slot) for item in matches))
        self.assertIn(reason, str(error.exception))
        return comparison


class ProjectionTests(WordStageTestCase):
    def test_selectable_counts_only_changed_slots_without_reasons(self):
        unchanged = self.compare(p('同じ'), p('同じ'))
        self.assertEqual((unchanged.differences, unchanged.selectable), (0, 0))
        before = make_docx(p('選べる文章の変更前') + p('注釈を付ける合成の文章の変更前') + p('末尾'))
        after = make_docx(p('選べる文章の変更後')
            + p(runs=run('注釈を付ける合成の文章の変更後') + '<w:r><w:commentReference w:id="1"/></w:r>') + p('末尾'),
            comments='<w:comment w:id="1">' + p('注記') + '</w:comment>')
        comparison = self.stage.compare(before, after)
        self.assertEqual(comparison.differences, 2)
        self.assertEqual(comparison.selectable, 1)
        self.assertEqual(len(comparison.unsupported), 1)

    def test_projections_have_matching_counts_ids_and_blank_padding(self):
        comparison = self.compare(p('保つ') + content_control(['甲', '乙']) + p('末尾'),
                                  p('保つ') + content_control(['甲の更新', '乙', '丙']) + p('追加') + p('末尾'))
        story = self.story(comparison)
        self.assertEqual(len(story.baseline.splitlines()), len(story.raw_changed.splitlines()))
        before, after = self.records(comparison, 'baseline'), self.records(comparison)
        self.assertEqual([r.id for r in before], [r.id for r in after])
        self.assertTrue(any(r.kind == 'blank' for r in before))
        self.assertTrue(any('/' in r.id for r in before))
        self.assertTrue(all(len(json.loads(line)) == 5 for line in story.raw_changed.splitlines()))

    def test_record_escaping_keeps_every_record_on_one_physical_line(self):
        text = 'タブ\t改行\n区切り\u0085中\u2028末\u2029'
        comparison = self.compare(p('元の文字'), p(text))
        line = next(line for line in self.story(comparison).raw_changed.splitlines() if text in self.stage.parse_record(line).text)
        self.assertIn('\\t', line)
        self.assertIn('\\n', line)
        for escape in ('\\u0085', '\\u2028', '\\u2029'):
            self.assertIn(escape, line)
        self.assertEqual(len(line.splitlines()), 1)
        self.assertEqual(self.stage.parse_record(line).text, text)

    def test_parse_record_rejects_malformed_records(self):
        for line in ('', '{', '{}', '[]', '["1","paragraph","","x"]',
                     '["1","paragraph","","x","",""]', '[1,"paragraph","","x",""]',
                     '["1","paragraph","",null,""]', '["1","paragraph","","x",false]'):
            with self.subTest(line=line), self.assertRaises(self.word.DocumentError):
                self.stage.parse_record(line)
        parsed = self.stage.parse_record('["12","blank","","",""]')
        self.assertEqual((parsed.id, parsed.kind, parsed.label, parsed.text, parsed.format),
                         ('12', 'blank', '', '', ''))

    def test_projection_is_identical_across_hash_seeds(self):
        code = (
            'import json,sys;sys.path[:0]=["scripts","tests"];'
            'import word_stage_ooxml as s;from word_fixtures import make_docx,p,table;'
            'a=make_docx(p("甲")+table([["一","二"]])+p("終わり"));'
            'b=make_docx(p("甲更新")+table([["一更新","二"]])+p("追記")+p("終わり"));'
            'c=s.compare(a,b);print(json.dumps([(x.id,x.baseline,x.changed,x.raw_changed) '
            'for x in c.stories],ensure_ascii=False))'
        )
        outputs = []
        for seed in ('1', '17', '321'):
            env = dict(os.environ, PYTHONHASHSEED=seed, PYTHONUTF8='1', PYTHONDONTWRITEBYTECODE='1')
            result = subprocess.run([sys.executable, '-B', '-c', code], cwd=ROOT, env=env,
                                    capture_output=True, encoding='utf-8', timeout=30)
            self.assertEqual(result.returncode, 0, result.stderr)
            outputs.append(result.stdout)
        self.assertEqual(outputs, [outputs[0]] * 3)


class ChooseTests(WordStageTestCase):
    def test_no_selection_and_single_line_selection(self):
        comparison = self.compare(p('変更前一') + p('変更前二'), p('変更後一') + p('変更後二'))
        story = self.story(comparison)
        self.assertEqual(comparison.choose('body', story.baseline.splitlines()), frozenset())
        selected = self.select(comparison, lambda record: record.text == '変更後二')
        self.assertEqual(len(selected), 1)
        self.assertIn('変更前一', self.word.render_text(comparison.rebuild(selected)))
        self.assertIn('変更後二', self.word.render_text(comparison.rebuild(selected)))

    def test_multi_line_unit_requires_every_changed_line(self):
        comparison = self.compare(content_control(['項目甲', '項目乙']) + p('末尾'),
                                  content_control(['項目甲の更新', '項目乙の更新', '追加項目']) + p('末尾'))
        with self.assertRaisesRegex(self.word.DocumentError, '選択が不完全'):
            self.select(comparison, lambda record: record.text == '項目甲の更新')
        selected = self.all_selected(comparison)
        self.assertEqual(len(selected), 1)
        text = self.word.render_text(comparison.rebuild(selected))
        self.assertIn('追加項目', text)
        self.assertNotIn('"項目乙"', text)

    def test_count_invented_line_and_line_from_another_position_are_rejected(self):
        comparison = self.compare(p('元の一') + p('元の二'), p('更新の一') + p('更新の二'))
        story = self.story(comparison)
        with self.assertRaisesRegex(self.word.DocumentError, '行数が変わっています'):
            comparison.choose('body', story.baseline.splitlines()[:-1])
        lines = list(story.baseline.splitlines())
        lines[0] = self.word.record_json(['1', 'paragraph', '', '捏造', ''])
        with self.assertRaisesRegex(self.word.DocumentError, '選択が不正'):
            comparison.choose('body', lines)
        lines[0] = story.raw_changed.splitlines()[1]
        with self.assertRaisesRegex(self.word.DocumentError, '選択が不正'):
            comparison.choose('body', lines)

    def test_merge_projection_cannot_select_blocked_slots(self):
        before = make_docx(p('コメントの前'))
        after = make_docx(p(runs=run('コメントの後') + '<w:r><w:commentReference w:id="1"/></w:r>'),
                          comments='<w:comment w:id="1">' + p('注釈') + '</w:comment>')
        comparison = self.stage.compare(before, after)
        story = self.story(comparison)
        self.assertTrue(comparison.unsupported)
        self.assertEqual(comparison.choose('body', story.changed.splitlines(), projection='changed'), frozenset())
        with self.assertRaisesRegex(self.word.DocumentError, 'コメント'):
            comparison.rebuild(comparison.choose('body', story.raw_changed.splitlines(), projection='raw_changed'))

    def test_rebuild_rejects_unknown_and_unchanged_keys(self):
        comparison = self.compare(p('同じ'), p('同じ'))
        for selection in (frozenset({('body', '999')}), frozenset({('body', '1')})):
            with self.subTest(selection=selection), self.assertRaisesRegex(self.word.DocumentError, 'Word の選択の行が不正'):
                comparison.rebuild(selection)


class ParagraphRebuildTests(WordStageTestCase):
    def test_replace_insert_delete_and_empty_selection(self):
        before = make_docx(p('冒頭') + p('変更前') + p('削除する文章') + p('末尾'))
        after = make_docx(p('冒頭') + p('変更後') + p('新しい文章') + p('末尾'))
        comparison = self.stage.compare(before, after)
        self.assertEqual(comparison.rebuild(frozenset()), before)
        result = comparison.rebuild(self.all_selected(comparison))
        self.assertEqual(self.word.render_text(result), self.word.render_text(after))
        # Distinct lengths force genuine insert/delete regions as well as replacement.
        insertion = make_docx(p('冒頭') + p('挿入した文章') + p('末尾'))
        self.assertEqual(self.word.render_text(self.rebuild_all(make_docx(p('冒頭') + p('末尾')), insertion)),
                         self.word.render_text(insertion))
        deletion = make_docx(p('冒頭') + p('末尾'))
        self.assertEqual(self.word.render_text(self.rebuild_all(before, deletion)), self.word.render_text(deletion))

    def test_untouched_paragraphs_parts_zip_metadata_and_order_are_preserved(self):
        options = dict(parts={'custom/untouched.bin': b'keep exact bytes'},
                       headers={'default': p('変更しないヘッダー')}, comment=b'Synthetic archive comment')
        before = make_docx(p('保持する段落', props='<w:spacing w:after="120"/>') + p('変更前') + p('末尾'), **options)
        after = make_docx(p('保持する段落', props='<w:spacing w:after="120"/>') + p('変更後') + p('末尾'), **options)
        result = self.rebuild_all(before, after)
        self.assertEqual(etree.tostring(self.tree(before).find('w:body/w:p', NS)),
                         etree.tostring(self.tree(result).find('w:body/w:p', NS)))
        with zipfile.ZipFile(io.BytesIO(before)) as original, zipfile.ZipFile(io.BytesIO(result)) as rebuilt:
            self.assertEqual(original.namelist(), rebuilt.namelist())
            self.assertEqual(original.comment, rebuilt.comment)
            for info, new_info in zip(original.infolist(), rebuilt.infolist()):
                self.assertEqual((info.filename, info.date_time, info.compress_type, info.external_attr, info.extra, info.comment),
                                 (new_info.filename, new_info.date_time, new_info.compress_type, new_info.external_attr, new_info.extra, new_info.comment))
                if info.filename != 'word/document.xml':
                    self.assertEqual(original.read(info), rebuilt.read(new_info), info.filename)


class CorruptionTests(WordStageTestCase):
    def test_serialized_corruptions_are_refused_before_returning(self):
        options = dict(rels=[('rImage', 'image', 'media/original.png'),
                             ('rUnused', 'hyperlink', 'https://example.invalid/keep', 'External')],
                       parts={'word/media/original.png': PNG, 'custom/untouched.bin': b'original'})
        before = make_docx(p('保持する文') + p('変更する前') + p(runs=image(doc_id=8)) + p('末尾'), **options)
        after = make_docx(p('保持する文') + p('変更した後') + p(runs=image(doc_id=8)) + p('末尾'), **options)
        comparison = self.stage.compare(before, after)
        selected = self.all_selected(comparison)
        writer = self.stage.write_zip

        def corrupt(kind, data):
            parts = unpack(data)
            if kind == 'untouched member':
                return change_parts(data, {'custom/untouched.bin': b'changed'})
            if kind == 'removed part':
                return change_parts(data, remove=('word/media/original.png',))
            if kind == 'dropped rel':
                root = etree.fromstring(parts['word/_rels/document.xml.rels'])
                root.remove(root[-1])
                return change_parts(data, {'word/_rels/document.xml.rels': xml_bytes(root)})
            root = etree.fromstring(parts['word/document.xml'])
            body = root.find('w:body', NS)
            paragraphs = body.findall('w:p', NS)
            if kind == 'selected skipped':
                paragraphs[1].find('.//w:t', NS).text = '変更する前'
            elif kind == 'unselected altered':
                paragraphs[0].find('.//w:t', NS).text = '勝手な変更'
            elif kind == 'extra paragraph':
                body.insert(2, etree.fromstring(xml_document('p', run('余分な段落')).encode()))
            elif kind == 'dangling numId':
                props = etree.SubElement(paragraphs[1], tag('pPr'))
                number = etree.SubElement(props, tag('numPr'))
                etree.SubElement(number, tag('numId')).set(tag('val'), '98765')
            elif kind == 'duplicate docPr':
                paragraphs[1].append(copy.deepcopy(paragraphs[2][0]))
            elif kind == 'broken bookmark':
                start = etree.SubElement(paragraphs[1], tag('bookmarkStart'))
                start.set(tag('id'), '98765')
                start.set(tag('name'), 'BrokenSyntheticBookmark')
            elif kind == 'undeclared Ignorable':
                root.set(tag('Ignorable', NS['mc']), 'missingSyntheticPrefix')
            return change_parts(data, {'word/document.xml': xml_bytes(root)})

        faults = ('selected skipped', 'unselected altered', 'extra paragraph', 'untouched member',
                  'dropped rel', 'dangling numId', 'duplicate docPr', 'broken bookmark',
                  'undeclared Ignorable', 'removed part')
        for fault in faults:
            with self.subTest(fault=fault):
                with mock.patch.object(self.stage, 'write_zip', side_effect=lambda *args, fault=fault, **kwargs: corrupt(fault, writer(*args, **kwargs))):
                    with self.assertRaises(self.word.DocumentError) as error:
                        comparison.rebuild(selected)
                self.assertEqual(str(error.exception), VALIDATION)


class ImageAndRelationshipTests(WordStageTestCase):
    def source_picture(self, **options):
        drawing = image().replace('<a:blip r:embed="rImage"/>',
            '<a:blip r:embed="rImage"><a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">'
            '<a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>'
            '</a:ext></a:extLst></a:blip>')
        return make_docx(p('保持') + p(runs=drawing) + p('末尾'),
                         rels=[('rImage', 'image', 'media/incoming.PNG')],
                         parts={'word/media/incoming.PNG': PNG}, **options)

    def test_picture_with_word_extension_is_copied_and_content_type_registered(self):
        source = self.source_picture()
        for mode in ('missing', 'conflicting'):
            with self.subTest(mode=mode):
                base = make_docx(p('保持') + p('画像に置換') + p('末尾'))
                parts = unpack(base)
                types = parts['[Content_Types].xml'].decode()
                types = types.replace('<Default Extension="png" ContentType="image/png"/>',
                                      '' if mode == 'missing' else '<Default Extension="png" ContentType="application/x-conflicting"/>')
                base = change_parts(base, {'[Content_Types].xml': types.encode()})
                comparison = self.stage.compare(base, source)
                self.assertEqual(comparison.unsupported, ())
                result = comparison.rebuild(self.all_selected(comparison))
                loaded = self.word.load(result)
                images = [name for name, value in loaded.parts.items() if name.startswith('word/media/') and value == PNG]
                self.assertEqual(len(images), 1)
                self.assertEqual(loaded.content_types.resolve(images[0]), 'image/png')
                types_root = self.tree(result, '[Content_Types].xml')
                if mode == 'missing':
                    self.assertEqual(types_root.xpath('ct:Default[@Extension="png"]/@ContentType', namespaces=NS), ['image/png'])
                else:
                    self.assertEqual(types_root.xpath('ct:Default[@Extension="png"]/@ContentType', namespaces=NS), ['application/x-conflicting'])
                    self.assertTrue(types_root.xpath('ct:Override[@ContentType="image/png"]', namespaces=NS))

    def test_identical_image_part_is_reused(self):
        base = make_docx(p('保持') + p('画像に置換') + p('末尾'), parts={'word/media/existing.png': PNG})
        result = self.rebuild_all(base, self.source_picture())
        self.assertEqual([name for name in unpack(result) if name.startswith('word/media/')], ['word/media/existing.png'])
        loaded = self.word.load(result)
        self.assertTrue(any(rel.path == 'word/media/existing.png' for rel in loaded.rels(loaded.main).values()))

    def test_svg_picture_copies_png_and_svg_relationships(self):
        drawing = image().replace('<a:blip r:embed="rImage"/>',
            '<a:blip r:embed="rImage"><a:extLst><a:ext uri="{SyntheticSvg}">'
            '<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rSvg"/>'
            '</a:ext></a:extLst></a:blip>')
        source = make_docx(p(runs=drawing) + p('末尾'),
                          rels=[('rImage', 'image', 'media/preview.png'), ('rSvg', 'image', 'media/vector.svg')],
                          parts={'word/media/preview.png': PNG, 'word/media/vector.svg': b'<svg xmlns="http://www.w3.org/2000/svg"/>'})
        parts = unpack(source)
        source = change_parts(source, {'[Content_Types].xml': parts['[Content_Types].xml'].replace(b'</Types>', b'<Default Extension="svg" ContentType="image/svg+xml"/></Types>')})
        result = self.rebuild_all(make_docx(p('画像前') + p('末尾')), source)
        loaded = self.word.load(result)
        image_rels = [rel for rel in loaded.rels(loaded.main).values() if rel.type.endswith('/image')]
        self.assertEqual(len(image_rels), 2)
        self.assertEqual({loaded.content_types.resolve(rel.path) for rel in image_rels}, {'image/png', 'image/svg+xml'})

    def test_header_image_creates_its_own_relationships_file(self):
        base = make_docx(p('本文'), headers={'default': p('前のヘッダー')})
        source = make_docx(p('本文'), headers={'default': p(runs=image())},
                          parts={'word/media/header.png': PNG,
                                 'word/_rels/header1.xml.rels': relationships([('rImage', 'image', 'media/header.png')])})
        comparison = self.stage.compare(base, source)
        selected = self.select(comparison, lambda record: record.kind == 'paragraph', 'header:1:default')
        result = comparison.rebuild(selected)
        loaded = self.word.load(result)
        self.assertIn('word/_rels/header1.xml.rels', loaded.parts)
        self.assertTrue(any(rel.type.endswith('/image') for rel in loaded.rels('word/header1.xml').values()))
        self.assertEqual(loaded.parts['word/document.xml'], unpack(base)['word/document.xml'])

    def test_external_hyperlink_is_reused_or_added(self):
        target = 'https://example.invalid/synthetic'
        source = make_docx(p(runs=hyperlink('リンクの更新')),
                          rels=[('rLink', 'hyperlink', target, 'External')])
        for existing in (False, True):
            with self.subTest(existing=existing):
                base = make_docx(p('元の文章'), rels=[('rId19', 'hyperlink', target, 'External')] if existing else [])
                loaded = self.word.load(self.rebuild_all(base, source))
                links = [rel for rel in loaded.rels(loaded.main).values() if rel.type.endswith('/hyperlink')]
                self.assertEqual(len(links), 1)
                self.assertEqual((links[0].target, links[0].external), (target, True))
                if existing:
                    self.assertEqual(links[0].id, 'rId19')


class StyleTests(WordStageTestCase):
    def test_style_names_map_japanese_and_english_ids(self):
        base = make_docx(p('見出し変更前', style='1'))
        source = make_docx(p('見出し変更後', style='Heading1'), japanese_ids=False)
        result = self.rebuild_all(base, source)
        self.assertEqual(self.tree(result).xpath('//w:pStyle/@w:val', namespaces=NS), ['1'])
        self.assertEqual(unpack(result)['word/styles.xml'], unpack(base)['word/styles.xml'])

    def test_missing_style_copies_dependency_chain_with_fresh_id_and_no_default(self):
        styles = ('<w:style w:type="paragraph" w:styleId="a" w:default="1" w:rsid="ABCDEF12">'
                  '<w:name w:val="Synthetic new style"/><w:basedOn w:val="Parent"/>'
                  '<w:next w:val="Next"/><w:link w:val="Character"/></w:style>'
                  '<w:style w:type="paragraph" w:styleId="Parent"><w:name w:val="Synthetic parent"/></w:style>'
                  '<w:style w:type="paragraph" w:styleId="Next"><w:name w:val="Synthetic next"/></w:style>'
                  '<w:style w:type="character" w:styleId="Character"><w:name w:val="Synthetic character"/></w:style>')
        base = make_docx(p('変更前'))
        source = make_docx(p('変更後', style='a'), styles=styles)
        result = self.rebuild_all(base, source)
        self.assertEqual(self.tree(result).xpath('//w:pStyle/@w:val', namespaces=NS), ['a_1'])
        root = self.tree(result, 'word/styles.xml')
        copied = root.xpath('w:style[w:name/@w:val="Synthetic new style"]', namespaces=NS)[0]
        self.assertNotIn(tag('default'), copied.attrib)
        self.assertNotIn(tag('rsid'), copied.attrib)
        for name in ('Parent', 'Next', 'Character'):
            self.assertTrue(root.xpath('w:style[@w:styleId="' + name + '"]', namespaces=NS))
        base_styles = self.tree(base, 'word/styles.xml')
        for old in base_styles:
            new = root.xpath('w:style[@w:styleId="' + old.get(tag('styleId')) + '"]', namespaces=NS)[0]
            self.assertEqual(etree.tostring(old, method='c14n'), etree.tostring(new, method='c14n'))

    def test_style_case_only_conflict_is_listed_before_rebuild(self):
        for base_name, source_name in (('Synthetic Case', 'synthetic case'),
                                       ('synthetic case', 'Synthetic Case')):
            with self.subTest(base_name=base_name, source_name=source_name):
                base = make_docx(p('合成の段落', style='Base'), styles=
                    '<w:style w:type="paragraph" w:styleId="Base"><w:name w:val="'
                    + base_name + '"/></w:style>')
                source = make_docx(p('合成の段落', style='Incoming'), styles=
                    '<w:style w:type="paragraph" w:styleId="Incoming"><w:name w:val="'
                    + source_name + '"/></w:style>')
                comparison = self.assert_blocked(base, source,
                    'スタイル名の大文字・小文字の違い「' + source_name + '」')
                self.assertEqual(comparison.differences, 1)
                self.assertEqual(comparison.selectable, 0)

    def test_style_type_conflict_and_missing_style_are_blocked(self):
        base = make_docx(p('前'), styles='<w:style w:type="character" w:styleId="Base"><w:name w:val="Synthetic conflict"/></w:style>')
        source = make_docx(p('後', style='Incoming'), styles='<w:style w:type="paragraph" w:styleId="Incoming"><w:name w:val="Synthetic conflict"/></w:style>')
        self.assert_blocked(base, source, 'スタイルの種類の違い「Synthetic conflict」')
        self.assert_blocked(make_docx(p('前')), make_docx(p('後', style='Missing')), '存在しないスタイルの参照')


class NumberingTests(WordStageTestCase):
    def test_equivalent_definition_reuses_base_num(self):
        base = make_docx(numbered('項目の前'), numbering=numbering_xml())
        source = make_docx(numbered('項目の後', 27), numbering=numbering_xml(27, 13, 'FFFFFFFF', 'EEEEEEEE'))
        result = self.rebuild_all(base, source)
        self.assertEqual(self.tree(result).xpath('//w:numId/@w:val', namespaces=NS), ['7'])
        self.assertEqual(unpack(result)['word/numbering.xml'], unpack(base)['word/numbering.xml'])

    def test_new_definition_schema_order_and_one_mapping_for_shared_list(self):
        base_numbering = numbering_xml().replace('</w:numbering>', '<w:numIdMacAtCleanup w:val="7"/></w:numbering>')
        base = make_docx(p('項目甲の前') + p('項目乙の前'), numbering=base_numbering)
        source = make_docx(numbered('項目甲の後', 27) + numbered('項目乙の後', 27), numbering=numbering_xml(27, 13, bullet=True))
        result = self.rebuild_all(base, source)
        ids = self.tree(result).xpath('//w:numId/@w:val', namespaces=NS)
        self.assertEqual(len(ids), 2)
        self.assertEqual(ids[0], ids[1])
        self.assertNotEqual(ids[0], '7')
        names = [etree.QName(node).localname for node in self.tree(result, 'word/numbering.xml')]
        self.assertEqual(names, ['abstractNum', 'abstractNum', 'num', 'num', 'numIdMacAtCleanup'])

    def test_missing_numbering_part_relationship_and_override_are_created(self):
        base = make_docx(p('リストの前'))
        source = make_docx(numbered('リストの後'), numbering=numbering_xml())
        result = self.rebuild_all(base, source)
        loaded = self.word.load(result)
        self.assertIn('word/numbering.xml', loaded.parts)
        self.assertTrue(any(rel.type.endswith('/numbering') and rel.path == 'word/numbering.xml' for rel in loaded.rels(loaded.main).values()))
        self.assertEqual(loaded.content_types.resolve('word/numbering.xml'),
                         'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml')

    def test_picture_bullet_list_style_and_undefined_num_are_blocked(self):
        base = make_docx(p('元の項目'))
        for fragment, reason in (('<w:lvlPicBulletId w:val="1"/>', '画像の行頭文字'),
                                 ('<w:numStyleLink w:val="SyntheticList"/>', 'リスト スタイル')):
            with self.subTest(reason=reason):
                number = numbering_xml().replace('<w:nsid', fragment + '<w:nsid', 1)
                self.assert_blocked(base, make_docx(numbered('新しい項目'), numbering=number), reason)
        self.assert_blocked(base, make_docx(numbered('新しい項目', 999)), '存在しない番号付けの参照')


class IdAndNamespaceTests(WordStageTestCase):
    def test_bookmark_sdt_revision_collisions_are_renumbered_and_paired(self):
        def decorated(text, ident):
            control = '<w:sdt><w:sdtPr><w:id w:val="' + str(ident) + '"/></w:sdtPr><w:sdtContent>' + run('制御') + '</w:sdtContent></w:sdt>'
            return p(runs=tracked(text, ident=ident) + control)
        before = make_docx(bookmark('保持', 'Original', 5) + decorated('保持する変更履歴', 7) + p('更新前') + p('末尾'))
        incoming = '<w:bookmarkStart w:id="5" w:name="Incoming"/>' + decorated('更新後', 7) + '<w:bookmarkEnd w:id="5"/>'
        after = make_docx(bookmark('保持', 'Original', 1) + decorated('保持する変更履歴', 2) + incoming + p('末尾'))
        result = self.rebuild_all(before, after)
        root = self.tree(result)
        starts = root.xpath('//w:bookmarkStart/@w:id', namespaces=NS)
        ends = root.xpath('//w:bookmarkEnd/@w:id', namespaces=NS)
        self.assertEqual(starts, ends)
        self.assertEqual(len(starts), len(set(starts)))
        self.assertEqual(len(root.xpath('//w:sdtPr/w:id/@w:val', namespaces=NS)),
                         len(set(root.xpath('//w:sdtPr/w:id/@w:val', namespaces=NS))))
        revisions = root.xpath('//w:ins/@w:id', namespaces=NS)
        self.assertEqual(len(revisions), len(set(revisions)))

    def test_paraids_are_stripped_and_source_namespaces_are_added_to_root(self):
        base = make_docx(p('そのまま') + p('変更前'))
        parts = unpack(base)
        xml = parts['word/document.xml'].decode()
        for prefix in ('w14', 'w16du'):
            xml = xml.replace(' xmlns:' + prefix + '="' + NS[prefix] + '"', '')
        base = change_parts(base, {'word/document.xml': xml.encode()})
        fragment = p(runs=tracked('変更後').replace('w:author=', 'w16du:dateUtc="2026-01-02T03:04:05Z" w:author=').replace('<w:r>', '<w:r><w:rPr><w14:ligatures w14:val="standard"/></w:rPr>'))
        fragment = fragment.replace('<w:p>', '<w:p w14:paraId="01234567" w14:textId="12345678">')
        source = make_docx(p('そのまま') + fragment)
        parts = unpack(source)
        source = change_parts(source, {'word/document.xml': parts['word/document.xml'].replace(b'<w:document ', b'<w:document mc:Ignorable="w14 w16du" ', 1)})
        result = self.rebuild_all(base, source)
        root = self.tree(result)
        self.assertEqual(root.nsmap['w14'], NS['w14'])
        self.assertEqual(root.nsmap['w16du'], NS['w16du'])
        self.assertIn('w14', root.get(tag('Ignorable', NS['mc']), '').split())
        self.assertIn('w16du', root.get(tag('Ignorable', NS['mc']), '').split())
        self.assertEqual(root.xpath('//@w14:paraId | //@w14:textId', namespaces=NS), [])
        before_paragraph = unpack(base)['word/document.xml'].split(b'<w:p>', 1)[1].split(b'</w:p>', 1)[0]
        after_paragraph = unpack(result)['word/document.xml'].split(b'<w:p>', 1)[1].split(b'</w:p>', 1)[0]
        self.assertEqual(before_paragraph, after_paragraph)

    def test_conflicting_namespace_prefix_is_blocked(self):
        base = make_docx(p('変更前'))
        parts = unpack(base)
        base = change_parts(base, {'word/document.xml': parts['word/document.xml'].replace(NS['w14'].encode(), b'urn:synthetic:conflicting-w14')})
        source = make_docx(p('変更後', props='<w14:glow w14:rad="100"/>'))
        self.assert_blocked(base, source, '名前空間の接頭辞の衝突 w14')

    def test_new_docpr_ids_are_unique_against_headers_and_notes(self):
        options = dict(headers={'default': p(runs=image(doc_id=30))},
                       footnotes='<w:footnote w:id="1">' + p(runs=image(doc_id=40)) + '</w:footnote>',
                       rels=[('rImage', 'image', 'media/pixel.png')], parts={'word/media/pixel.png': PNG})
        before = make_docx(p('画像を入れる場所') + p(runs=image(doc_id=8)) + p('末尾'), **options)
        after = make_docx(p(runs=image(doc_id=8)) + p(runs=image(doc_id=8)) + p('末尾'), **options)
        comparison = self.stage.compare(before, after)
        result = comparison.rebuild(self.all_selected(comparison))
        ids = self.tree(result).xpath('//wp:docPr/@id', namespaces=NS)
        self.assertEqual(len(ids), 2)
        self.assertEqual(ids[1], '8')
        self.assertGreater(int(ids[0]), 40)
        self.assertEqual(unpack(result)['word/header1.xml'], unpack(before)['word/header1.xml'])
        self.assertEqual(unpack(result)['word/footnotes.xml'], unpack(before)['word/footnotes.xml'])

    def test_preexisting_docpr_duplicates_are_tolerated(self):
        options = dict(rels=[('rImage', 'image', 'media/pixel.png')], parts={'word/media/pixel.png': PNG})
        before = make_docx(p(runs=image(doc_id=8)) + p(runs=image(doc_id=8)) + p('更新前'), **options)
        after = make_docx(p(runs=image(doc_id=8)) + p(runs=image(doc_id=8)) + p('更新後'), **options)
        result = self.rebuild_all(before, after)
        self.assertEqual(self.tree(result).xpath('//wp:docPr/@id', namespaces=NS), ['8', '8'])


class SectionTests(WordStageTestCase):
    def test_section_properties_keep_base_header_references(self):
        base = make_docx(p('本文'), headers={'default': ('word/header1.xml', p('保持するヘッダー'))})
        source = make_docx(p('本文'), headers={'default': ('word/header3.xml', p('取り込まないヘッダー'))})
        parts = unpack(source)
        source = change_parts(source, {'word/document.xml': parts['word/document.xml'].replace(b'w:w="11906" w:h="16838"', b'w:w="16838" w:h="11906" w:orient="landscape"')})
        comparison = self.stage.compare(base, source)
        result = comparison.rebuild(self.select(comparison, lambda record: record.kind == 'section'))
        self.assertEqual(self.tree(result).xpath('//w:sectPr/w:headerReference/@r:id', namespaces=NS),
                         self.tree(base).xpath('//w:sectPr/w:headerReference/@r:id', namespaces=NS))
        self.assertEqual(unpack(result)['word/header1.xml'], unpack(base)['word/header1.xml'])
        self.assertEqual(self.tree(result).xpath('//w:sectPr/w:pgSz/@w:orient', namespaces=NS), ['landscape'])

    def test_section_additions_and_removals_are_blocked(self):
        one = make_docx(p('段落甲') + p('段落乙'))
        two = make_docx(p('段落甲', props='<w:sectPr><w:type w:val="continuous"/></w:sectPr>') + p('段落乙'))
        self.assert_blocked(one, two, 'セクション区切りの追加・削除')
        self.assert_blocked(two, one, 'セクション区切りの追加・削除')

    def test_deleted_section_host_reattaches_to_previous_paragraph(self):
        section = '<w:sectPr><w:type w:val="continuous"/></w:sectPr>'
        base = make_docx(p('先頭') + p('削除する段落', props=section) + p('末尾'))
        source = make_docx(p('先頭', props=section) + p('末尾'))
        comparison = self.stage.compare(base, source)
        result = comparison.rebuild(self.all_selected(comparison))
        root = self.tree(result)
        self.assertEqual(root.xpath('w:body/w:p[1]/w:pPr/w:sectPr/w:type/@w:val', namespaces=NS), ['continuous'])
        self.assertEqual(root.xpath('w:body/w:p/w:r/w:t/text()', namespaces=NS), ['先頭', '末尾'])

    def test_consecutive_sections_and_a_body_ending_in_table_are_refused(self):
        section = '<w:sectPr><w:type w:val="continuous"/></w:sectPr>'
        base = make_docx(p('先頭', props=section) + p('削除する段落', props=section) + p('末尾'))
        source = make_docx(p('先頭', props=section) + p('末尾'))
        comparison = self.stage.compare(base, source)
        deleted_ids = {record.id for record in self.records(comparison, 'baseline') if record.text == '削除する段落'}
        selected = self.select(comparison, lambda record: record.kind == 'blank' and record.id in deleted_ids)
        with self.assertRaisesRegex(self.word.DocumentError, 'セクション区切りの前に段落が必要'):
            comparison.rebuild(selected)
        comparison = self.stage.compare(make_docx(table([['値']]) + p('末尾')), make_docx(table([['値']])))
        with self.assertRaisesRegex(self.word.DocumentError, '最後には段落が必要'):
            comparison.rebuild(self.select(comparison, lambda record: record.kind == 'blank'))


class TableTests(WordStageTestCase):
    def test_cell_selection_keeps_widths_and_other_cells(self):
        base = make_docx(table([['甲', '乙']], widths=[1000, 2000]) + p('末尾'))
        source = make_docx(table([['甲の更新', '乙の更新']], widths=[3000, 4000],
                                 cell_props={(0, 0): '<w:shd w:fill="FFFF00"/>'}) + p('末尾'))
        comparison = self.stage.compare(base, source)
        selected = self.select(comparison, lambda record: record.kind == 'cell' and record.label == 'R1C1')
        self.assertEqual(len(selected), 1)
        result = comparison.rebuild(selected)
        root = self.tree(result)
        self.assertEqual(root.xpath('//w:tcW/@w:w', namespaces=NS), ['1000', '2000'])
        self.assertEqual(root.xpath('//w:tc//w:t/text()', namespaces=NS), ['甲の更新', '乙'])
        self.assertEqual(root.xpath('//w:tc[1]/w:tcPr/w:shd/@w:fill', namespaces=NS), ['FFFF00'])
        self.assertEqual(etree.tostring(self.tree(base).find('.//w:tc[2]', NS)),
                         etree.tostring(root.find('.//w:tc[2]', NS)))

    def test_table_geometry_and_row_properties_have_independent_slots(self):
        base = make_docx(table([['値']], widths=[1000]) + p('末尾'))
        source = make_docx(table([['更新値']], widths=[3000], row_props={0: '<w:cantSplit/>'}) + p('末尾'))
        comparison = self.stage.compare(base, source)
        geometry = comparison.rebuild(self.select(comparison, lambda record: record.kind == 'table'))
        root = self.tree(geometry)
        self.assertEqual(root.xpath('//w:gridCol/@w:w | //w:tcW/@w:w', namespaces=NS), ['3000', '3000'])
        self.assertEqual(root.xpath('//w:tc//w:t/text()', namespaces=NS), ['値'])
        self.assertFalse(root.xpath('//w:cantSplit', namespaces=NS))
        row = comparison.rebuild(self.select(comparison, lambda record: record.kind == 'row'))
        root = self.tree(row)
        self.assertTrue(root.xpath('//w:cantSplit', namespaces=NS))
        self.assertEqual(root.xpath('//w:tcW/@w:w', namespaces=NS), ['1000'])
        self.assertEqual(root.xpath('//w:tc//w:t/text()', namespaces=NS), ['値'])

    def test_row_addition_and_merge_changes_form_whole_table_slots(self):
        base = make_docx(table([['値甲', '値乙']]) + p('末尾'))
        for changed in (table([['値甲', '値乙'], ['追加甲', '追加乙']]),
                        table([['値甲の更新', '値乙']], widths=[3000, 4000], cell_props={(0, 0): '<w:gridSpan w:val="2"/>'}),
                        table([['値甲の更新', '値乙']], widths=[3000, 4000], cell_props={(0, 0): '<w:vMerge w:val="restart"/>'})):
            with self.subTest(changed=changed):
                comparison = self.stage.compare(base, make_docx(changed + p('末尾')))
                table_records = [record for record in self.records(comparison) if record.kind in ('table', 'row', 'cell')]
                self.assertTrue(all('/' in record.id for record in table_records))
                self.assertEqual(len(self.all_selected(comparison)), 1)
                with self.assertRaisesRegex(self.word.DocumentError, '選択が不完全'):
                    self.select(comparison, lambda record: record.kind == 'cell')

    def test_unsupported_cell_does_not_block_another_sliced_cell(self):
        commented = p(runs=run('注釈付きの旧値') + '<w:r><w:commentReference w:id="1"/></w:r>')
        updated = p(runs=run('注釈付きの新値') + '<w:r><w:commentReference w:id="1"/></w:r>')
        options = {'comments': '<w:comment w:id="1">' + p('注記') + '</w:comment>'}
        before = make_docx(table([[commented, '旧値']]) + p('末尾'), **options)
        after = make_docx(table([[updated, '新値']]) + p('末尾'), **options)
        comparison = self.stage.compare(before, after)
        blocked = [item for item in comparison.unsupported if 'コメント' in item.reason]
        self.assertEqual(len(blocked), 1)
        self.assertTrue(blocked[0].slot.endswith('R1C1'))
        selected = self.select(comparison, lambda record: record.kind == 'cell' and record.label == 'R1C2')
        result = comparison.rebuild(selected)
        self.assertEqual(self.tree(result).xpath('//w:tc//w:t/text()', namespaces=NS), ['注釈付きの旧値', '新値'])


class HeaderFooterTests(WordStageTestCase):
    def test_pairing_uses_roles_when_header_and_footer_parts_are_renamed(self):
        base = make_docx(p('本文'), headers={'default': ('word/header1.xml', p('前のヘッダー'))},
                         footers={'default': ('word/footer1.xml', p('前のフッター'))})
        source = make_docx(p('本文'), headers={'default': ('word/header3.xml', p('後のヘッダー'))},
                           footers={'default': ('word/footer3.xml', p('後のフッター'))})
        comparison = self.stage.compare(base, source)
        self.assertEqual([story.id for story in comparison.stories], ['body', 'header:1:default', 'footer:1:default'])
        result = comparison.rebuild(self.all_selected(comparison))
        parts = unpack(result)
        self.assertIn('後のヘッダー'.encode(), parts['word/header1.xml'])
        self.assertIn('後のフッター'.encode(), parts['word/footer1.xml'])
        self.assertNotIn('word/header3.xml', parts)
        self.assertNotIn('word/footer3.xml', parts)

    def test_unpaired_header_is_listed_as_unsupported(self):
        comparison = self.stage.compare(make_docx(p('本文')), make_docx(p('本文'), headers={'first': p('新しいヘッダー')}))
        self.assertTrue(any('ヘッダー・フッターの追加・削除' in item.reason for item in comparison.unsupported))
        self.assertEqual(comparison.selectable, 0)


class ReferenceTests(WordStageTestCase):
    def pair(self, heading_changed=True):
        before = make_docx(p(runs=hyperlink('旧参照', anchor='_TocOld')) + bookmark('見出し', '_TocOld', 1) + p('末尾'))
        after = make_docx(p(runs=hyperlink('新参照', anchor='_TocNew')) + bookmark('見出しの更新' if heading_changed else '見出し', '_TocNew', 4) + p('末尾'))
        return self.stage.compare(before, after)

    def test_heading_without_kept_reference_target_and_reference_without_heading_fail(self):
        comparison = self.pair()
        with self.assertRaisesRegex(self.word.DocumentError, '参照先「_TocOld」が見つからなくなります'):
            comparison.rebuild(self.select(comparison, lambda record: record.text == '見出しの更新'))
        with self.assertRaisesRegex(self.word.DocumentError, '参照先「_TocNew」が見つからなくなります'):
            comparison.rebuild(self.select(comparison, lambda record: record.text == '新参照'))

    def test_equal_heading_is_swapped_for_selected_updated_reference(self):
        comparison = self.pair(heading_changed=False)
        result = comparison.rebuild(self.select(comparison, lambda record: record.text == '新参照'))
        self.assertEqual(comparison.last_swaps, 1)
        self.assertEqual(self.tree(result).xpath('//w:bookmarkStart/@w:name', namespaces=NS), ['_TocNew'])
        self.assertIn('見出し', self.word.render_text(result))

    def test_equal_swap_does_not_break_a_remaining_base_reference(self):
        before = make_docx(p(runs=hyperlink('更新する参照', anchor='_TocOld'))
                           + p(runs=hyperlink('保持する参照', anchor='_TocOld')) + bookmark('見出し', '_TocOld', 1) + p('末尾'))
        after = make_docx(p(runs=hyperlink('参照の更新後', anchor='_TocNew'))
                          + p(runs=hyperlink('保持する参照', anchor='_TocNew')) + bookmark('見出し', '_TocNew', 2) + p('末尾'))
        comparison = self.stage.compare(before, after)
        with self.assertRaisesRegex(self.word.DocumentError, '参照先「_TocOld」が見つからなくなります'):
            comparison.rebuild(self.select(comparison, lambda record: record.text == '参照の更新後'))

    def test_visible_bookmark_duplicate_is_refused(self):
        before = make_docx(bookmark('保持', 'VisibleTarget', 1) + p('更新前') + p('末尾'))
        after = make_docx(bookmark('保持の更新', 'OtherTarget', 1) + bookmark('更新後', 'VisibleTarget', 2) + p('末尾'))
        comparison = self.stage.compare(before, after)
        with self.assertRaisesRegex(self.word.DocumentError, 'ブックマーク「VisibleTarget」が重複します'):
            comparison.rebuild(self.select(comparison, lambda record: record.text == '更新後'))

    def test_hidden_duplicate_and_orphaned_end_are_dropped(self):
        before = make_docx(bookmark('保持', '_HiddenTarget', 1) + p('更新前') + p('末尾'))
        after = make_docx(bookmark('保持', '_OtherTarget', 1) + bookmark('更新後', '_HiddenTarget', 2) + p('末尾'))
        comparison = self.stage.compare(before, after)
        result = comparison.rebuild(self.select(comparison, lambda record: record.text == '更新後'))
        self.assertEqual(self.tree(result).xpath('//w:bookmarkStart/@w:name', namespaces=NS), ['_HiddenTarget'])
        before = make_docx('<w:bookmarkStart w:id="1" w:name="_HlkSpan"/>' + p('削除する') + p('保持する') + '<w:bookmarkEnd w:id="1"/>' + p('末尾'))
        after = make_docx(p('保持する') + p('末尾'))
        comparison = self.stage.compare(before, after)
        result = comparison.rebuild(self.all_selected(comparison))
        self.assertEqual(self.tree(result).xpath('//w:bookmarkStart | //w:bookmarkEnd', namespaces=NS), [])

    def spanning_pair(self, changed_end=False, blocked_start=False):
        def body(name, ident, reference, updated):
            first = p(runs=run('見出しの先頭') + ('<w:r><w:commentReference w:id="1"/></w:r>' if blocked_start else ''))
            second = p('見出しの末尾の更新' if updated and changed_end else '見出しの末尾')
            return (p(runs=hyperlink(reference, anchor=name)) + '<w:bookmarkStart w:id="' + str(ident) + '" w:name="' + name + '"/>'
                    + first + second + '<w:bookmarkEnd w:id="' + str(ident) + '"/>' + p('末尾'))
        options = {'comments': '<w:comment w:id="1">' + p('固定の注釈') + '</w:comment>'} if blocked_start else {}
        return self.stage.compare(make_docx(body('_TocOld', 1, '旧参照', False), **options),
                                  make_docx(body('_TocNew', 8, '新参照', True), **options))

    def test_equal_hidden_range_swaps_both_endpoints(self):
        comparison = self.spanning_pair()
        result = comparison.rebuild(self.select(comparison, lambda record: record.text == '新参照'))
        self.assertEqual(comparison.last_swaps, 2)
        root = self.tree(result)
        self.assertEqual(root.xpath('//w:bookmarkStart/@w:name', namespaces=NS), ['_TocNew'])
        self.assertEqual(root.xpath('//w:bookmarkStart/@w:id', namespaces=NS),
                         root.xpath('//w:bookmarkEnd/@w:id', namespaces=NS))

    def test_equal_hidden_range_cannot_swap_a_changed_unselected_endpoint(self):
        comparison = self.spanning_pair(changed_end=True)
        with self.assertRaisesRegex(self.word.DocumentError, '参照先「_TocNew」が見つからなくなります'):
            comparison.rebuild(self.select(comparison, lambda record: record.text == '新参照'))

    def test_equal_hidden_target_with_intrinsic_reason_cannot_be_swapped(self):
        comparison = self.spanning_pair(blocked_start=True)
        with self.assertRaisesRegex(self.word.DocumentError, '参照先「_TocNew」が見つからなくなります'):
            comparison.rebuild(self.select(comparison, lambda record: record.text == '新参照'))

    def test_visible_bookmark_duplicate_between_stories_has_specific_error(self):
        before = make_docx(bookmark('本文の保持', 'VisibleTarget', 1), headers={'default': p('ヘッダーの前')})
        after = make_docx(bookmark('本文の変更', 'OtherTarget', 1),
                          headers={'default': bookmark('ヘッダーの後', 'VisibleTarget', 2)})
        comparison = self.stage.compare(before, after)
        selected = self.select(comparison, lambda record: record.text == 'ヘッダーの後', 'header:1:default')
        with self.assertRaisesRegex(self.word.DocumentError, 'ブックマーク「VisibleTarget」が重複します'):
            comparison.rebuild(selected)

    def test_equal_swap_preserves_references_in_untouched_notes(self):
        before = make_docx(p(runs=hyperlink('旧参照', anchor='_TocOld')) + bookmark('見出し', '_TocOld', 1),
                           footnotes='<w:footnote w:id="1">' + p(runs=hyperlink('脚注の参照', anchor='_TocOld')) + '</w:footnote>')
        after = make_docx(p(runs=hyperlink('新参照', anchor='_TocNew')) + bookmark('見出し', '_TocNew', 2),
                          footnotes='<w:footnote w:id="1">' + p(runs=hyperlink('脚注の参照', anchor='_TocNew')) + '</w:footnote>')
        comparison = self.stage.compare(before, after)
        with self.assertRaisesRegex(self.word.DocumentError, '参照先「_TocOld」が見つからなくなります'):
            comparison.rebuild(self.select(comparison, lambda record: record.text == '新参照'))

    def test_equal_target_in_sliced_cell_swaps_only_that_piece(self):
        commented = p(runs=run('注釈付きの固定セル') + '<w:r><w:commentReference w:id="1"/></w:r>')
        options = {'comments': '<w:comment w:id="1">' + p('注記') + '</w:comment>'}
        before = make_docx(p(runs=hyperlink('旧参照', anchor='_TocOld'))
                           + table([[p(runs='<w:bookmarkStart w:id="1" w:name="_TocOld"/>' + run('見出し') + '<w:bookmarkEnd w:id="1"/>'), commented]]) + p('末尾'), **options)
        after = make_docx(p(runs=hyperlink('新参照', anchor='_TocNew'))
                          + table([[p(runs='<w:bookmarkStart w:id="8" w:name="_TocNew"/>' + run('見出し') + '<w:bookmarkEnd w:id="8"/>'), commented]]) + p('末尾'), **options)
        comparison = self.stage.compare(before, after)
        result = comparison.rebuild(self.select(comparison, lambda record: record.text == '新参照'))
        self.assertEqual(comparison.last_swaps, 1)
        root = self.tree(result)
        self.assertEqual(root.xpath('//w:bookmarkStart/@w:name', namespaces=NS), ['_TocNew'])
        self.assertEqual(etree.tostring(self.tree(before).find('.//w:tc[2]', NS)),
                         etree.tostring(root.find('.//w:tc[2]', NS)))


class UnsupportedTests(WordStageTestCase):
    def test_intrinsic_reasons_are_listed_hidden_and_rejected(self):
        cases = (
            ('コメント', p(runs=run('後') + '<w:r><w:commentReference w:id="1"/></w:r>')),
            ('脚注・文末脚注', p(runs=run('後') + '<w:r><w:footnoteReference w:id="1"/></w:r>')),
            ('グラフ', p(runs=run('後') + '<w:r><w:drawing><c:chart/></w:drawing></w:r>')),
            ('SmartArt', p(runs=run('後') + '<w:r><d:relIds xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></w:r>')),
            ('図形・テキストボックス', p(runs=run('後') + '<w:r><wps:wsp/></w:r>')),
            ('VML の図形・画像', p(runs=run('後') + '<w:r><w:pict><v:shape/></w:pict></w:r>')),
            ('埋め込みオブジェクト（OLE）', p(runs=run('後') + '<w:r><w:object/></w:r>')),
            ('ActiveX コントロール', p(runs=run('後') + '<w:r><w:control/></w:r>')),
            ('インク・3D モデル', p(runs=run('後') + '<w:r><w:contentPart/></w:r>')),
            ('変更履歴の移動', p(runs=tracked('後', kind='moveTo'))),
            ('データにリンクしたコンテンツ コントロール', '<w:sdt><w:sdtPr><w:dataBinding w:xpath="/Synthetic"/></w:sdtPr><w:sdtContent>' + p('後') + '</w:sdtContent></w:sdt>'),
            ('繰り返しセクション', '<w:sdt><w:sdtPr><w15:repeatingSection xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"/></w:sdtPr><w:sdtContent>' + p('後') + '</w:sdtContent></w:sdt>'),
            ('プレースホルダー表示中のコンテンツ コントロール', '<w:sdt><w:sdtPr><w:showingPlcHdr/><w:placeholder><w:docPart w:val="MissingSyntheticPart"/></w:placeholder></w:sdtPr><w:sdtContent>' + p('後') + '</w:sdtContent></w:sdt>'),
            ('埋め込み文書（altChunk）', '<w:altChunk/>'),
            ('サブ文書', '<w:subDoc/>'),
            ('カスタム XML', '<w:customXml>' + p('後') + '</w:customXml>'),
            ('未対応の要素 <x:unknown>', p(runs=run('後') + '<w:r><x:unknown xmlns:x="urn:synthetic:unknown"/></w:r>')),
            ('閉じていないフィールド', p(runs='<w:r><w:fldChar w:fldCharType="begin"/><w:instrText> TOC </w:instrText></w:r>' + run('後'))),
            ('セクション区切りを含むまとまり', table([[p('後', props='<w:sectPr/>')]])),
        )
        for reason, body in cases:
            with self.subTest(reason=reason):
                self.assert_blocked(make_docx(p('前') + p('末尾')), make_docx(body + p('末尾')), reason)

    def test_base_only_unsupported_presence_also_blocks_the_slot(self):
        before = make_docx(p(runs=run('更新前') + '<w:r><w:commentReference w:id="1"/></w:r>') + p('末尾'),
                           comments='<w:comment w:id="1">' + p('注記') + '</w:comment>')
        self.assert_blocked(before, make_docx(p('更新後') + p('末尾')), 'コメント')

    def test_planner_duplicate_style_names_and_unsupported_rels(self):
        styles = default_styles().replace('</w:styles>',
            '<w:style w:type="paragraph" w:styleId="DupA"><w:name w:val="Duplicate"/></w:style>'
            '<w:style w:type="paragraph" w:styleId="DupB"><w:name w:val="Duplicate"/></w:style></w:styles>')
        self.assert_blocked(make_docx(p('前')), make_docx(p('後', style='DupA'), styles=styles), 'スタイル名の重複「Duplicate」')
        source = make_docx(p(runs=run('後') + '<w:r><w:drawing><a:blip r:embed="rUnsupported"/></w:drawing></w:r>'),
                          rels=[('rUnsupported', 'hdphoto', 'media/photo.wdp')], parts={'word/media/photo.wdp': b'synthetic'})
        self.assert_blocked(make_docx(p('前')), source, '未対応の参照 hdphoto')

    def test_header_role_change_is_blocked(self):
        base = make_docx(p('本文'), headers={'default': p('同じ')})
        source = make_docx(p('本文'), headers={'first': p('同じ')})
        comparison = self.stage.compare(base, source)
        self.assertTrue(any('ヘッダー・フッターの構成' in item.reason for item in comparison.unsupported))


class PerformanceTests(WordStageTestCase):
    def test_five_thousand_paragraphs_with_empty_lines_prepare_quickly(self):
        fragments = [p('' if index % 5 == 0 else '試験段落 ' + str(index)) for index in range(5000)]
        before = make_docx(''.join(fragments))
        fragments[2311] = p('試験段落 2311 の変更')
        fragments[4703] = p('試験段落 4703 の変更')
        after = make_docx(''.join(fragments))
        started = time.monotonic()
        comparison = self.stage.compare(before, after)
        elapsed = time.monotonic() - started
        self.assertEqual(comparison.differences, 2)
        self.assertLess(elapsed, 15, 'prepare exceeded the capped alignment budget')

    def test_completely_different_two_thousand_paragraph_regions_are_capped(self):
        before = make_docx(''.join(p('甲' * 20 + str(index)) for index in range(2000)))
        after = make_docx(''.join(p('乙' * 20 + str(index)) for index in range(2000)))
        started = time.monotonic()
        comparison = self.stage.compare(before, after)
        self.assertLess(time.monotonic() - started, 15)
        self.assertEqual(comparison.differences, 2000)
        self.assertEqual(len(self.story(comparison).baseline.splitlines()), 2001)


class PlaceholderDestinationTests(WordStageTestCase):
    def test_placeholder_uses_destination_glossary_even_when_source_lacks_it(self):
        def control(text):
            return ('<w:sdt><w:sdtPr><w:showingPlcHdr/><w:placeholder>'
                    '<w:docPart w:val="SyntheticDestinationPlaceholder"/>'
                    '</w:placeholder></w:sdtPr><w:sdtContent>' + p(text)
                    + '</w:sdtContent></w:sdt>')
        glossary = xml_document('glossaryDocument',
            '<w:docParts><w:docPart><w:docPartPr>'
            '<w:name w:val="SyntheticDestinationPlaceholder"/>'
            '</w:docPartPr><w:docPartBody>' + p('合成の登録済み部品')
            + '</w:docPartBody></w:docPart></w:docParts>')
        before = make_docx(control('部品の前') + p('末尾'),
            rels=[('rGlossary', 'glossaryDocument', 'glossary/document.xml')],
            parts={'word/glossary/document.xml': glossary})
        after = make_docx(control('部品の後') + p('末尾'))
        comparison = self.stage.compare(before, after)
        self.assertEqual(comparison.unsupported, ())
        result = comparison.rebuild(self.all_selected(comparison))
        self.assertEqual(unpack(result)['word/glossary/document.xml'],
                         unpack(before)['word/glossary/document.xml'])
        text = self.word.render_text(result)
        self.assertIn('部品の後', text)
        self.assertNotIn('部品の前', text)


class DependencyValidationTests(WordStageTestCase):
    def test_untouched_styles_cannot_reference_missing_numbering(self):
        styles = default_styles().replace('</w:styles>',
            '<w:style w:type="paragraph" w:styleId="UnreferencedBroken">'
            '<w:name w:val="Synthetic broken list style"/><w:pPr><w:numPr>'
            '<w:numId w:val="999"/></w:numPr></w:pPr></w:style></w:styles>')
        comparison = self.stage.compare(make_docx(p('本文の前'), styles=styles),
                                        make_docx(p('本文の後'), styles=styles))
        with self.assertRaisesRegex(self.word.DocumentError, '^' + VALIDATION + '$'):
            comparison.rebuild(self.all_selected(comparison))

    def test_copied_style_relationships_are_planned_with_the_style_owner(self):
        styles = default_styles().replace('</w:styles>',
            '<w:style w:type="paragraph" w:styleId="SyntheticLinkedStyle">'
            '<w:name w:val="Synthetic linked style"/><w:rPr>'
            '<w:rFonts r:id="rUnsupported"/></w:rPr></w:style></w:styles>')
        source = make_docx(p('更新後', style='SyntheticLinkedStyle'), styles=styles,
            parts={'word/_rels/styles.xml.rels': relationships([
                ('rUnsupported', 'printerSettings', 'printerSettings/synthetic.bin')]),
                'word/printerSettings/synthetic.bin': b'synthetic'})
        self.assert_blocked(make_docx(p('更新前')), source, '未対応の参照 printerSettings')

    def test_record_id_grammar_kind_and_blank_payload_are_checked(self):
        for identifier in ('0', '01', '-1', '1/0', '1/01', '1 R0', '1 R01', '1 R1C0', '1 R1C01', '1 R1C1/2'):
            with self.subTest(identifier=identifier), self.assertRaises(self.word.DocumentError):
                self.stage.parse_record(self.word.record_json([identifier, 'paragraph', '', '', '']))
        for fields in (['1', 'unknown', '', '', ''], ['1', 'blank', '', 'payload', '']):
            with self.subTest(fields=fields), self.assertRaises(self.word.DocumentError):
                self.stage.parse_record(self.word.record_json(fields))
        for identifier in ('1', '12/3', '4 R2', '4 R2C3'):
            self.assertEqual(self.stage.parse_record(self.word.record_json([identifier, 'blank', '', '', ''])).id,
                             identifier)


if __name__ == '__main__':
    unittest.main()
