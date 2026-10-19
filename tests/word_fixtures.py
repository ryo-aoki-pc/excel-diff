"""Synthetic OOXML package builders shared by the Word tests.

All document text and identifiers are invented. These helpers intentionally use
XML strings rather than a Word library so tests can exercise save-time noise and
unusual, but legal, OOXML directly.
"""

import base64
import io
import re
import zipfile
from xml.sax.saxutils import escape, quoteattr


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
STRICT_W = "http://purl.oclc.org/ooxml/wordprocessingml/main"
STRICT_R = "http://purl.oclc.org/ooxml/officeDocument/relationships"
MAIN_TYPES = {
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    "docm": "application/vnd.ms-word.document.macroEnabled.main+xml",
    "dotx": "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
    "dotm": "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
}
NAMESPACES = {
    "w": W, "r": R,
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w16du": "http://schemas.microsoft.com/office/word/2023/wordml/word16du",
    "w16se": "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "a16": "http://schemas.microsoft.com/office/drawing/2014/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "v": "urn:schemas-microsoft-com:vml", "o": "urn:schemas-microsoft-com:office:office",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
}
NS_DECL = " ".join("xmlns:" + name + "=" + quoteattr(uri) for name, uri in NAMESPACES.items())
PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
)


def xml_document(name, content, attrs=""):
    """Wrap a fragment in a standalone Word root with useful namespaces."""
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:' + name + ' ' + NS_DECL + (' ' + attrs if attrs else '') + '>'
            + content + '</w:' + name + '>')


def run(text, props="", preserve=True):
    """Return a run, escaping text while accepting literal rPr XML."""
    return ('<w:r>' + ('<w:rPr>' + props + '</w:rPr>' if props else '')
            + '<w:t' + (' xml:space="preserve"' if preserve else '') + '>'
            + escape(text) + '</w:t></w:r>')


def p(text="", style=None, props="", runs=None):
    """Return a paragraph with optional direct properties or literal runs."""
    properties = ('<w:pStyle w:val=' + quoteattr(style) + '/>' if style else '') + props
    return ('<w:p>' + ('<w:pPr>' + properties + '</w:pPr>' if properties else '')
            + (run(text) if runs is None else runs) + '</w:p>')


def table(rows, *, props="", row_props=None, cell_props=None, widths=None):
    """Build a rectangular table; cells accept text or literal block XML."""
    count = max((len(row) for row in rows), default=0)
    widths = widths or [2000] * count
    result = '<w:tbl>' + ('<w:tblPr>' + props + '</w:tblPr>' if props else '')
    result += '<w:tblGrid>' + ''.join('<w:gridCol w:w="' + str(w) + '"/>' for w in widths) + '</w:tblGrid>'
    for ri, row in enumerate(rows):
        result += '<w:tr>'
        rp = (row_props or {}).get(ri, '')
        result += '<w:trPr>' + rp + '</w:trPr>' if rp else ''
        for ci, text in enumerate(row):
            cp = '<w:tcW w:w="' + str(widths[ci]) + '" w:type="dxa"/>' + (cell_props or {}).get((ri, ci), '')
            result += '<w:tc><w:tcPr>' + cp + '</w:tcPr>'
            result += text if text.startswith('<') else p(text)
            result += '</w:tc>'
        result += '</w:tr>'
    return result + '</w:tbl>'


def default_styles(japanese_ids=True):
    """Use stable names under Japanese or English Word-style identifiers."""
    ids = ('a', 'a0', 'a1', '1', '2', 'SyntheticJaList') if japanese_ids else (
        'Normal', 'DefaultParagraphFont', 'TableNormal', 'Heading1', 'Heading2', 'ListParagraph')
    normal, char, grid, heading1, heading2, listing = ids
    definitions = (
        ('paragraph', normal, 'Normal', True),
        ('character', char, 'Default Paragraph Font', True),
        ('table', grid, 'Normal Table', True),
        ('paragraph', heading1, 'heading 1', False),
        ('paragraph', heading2, 'heading 2', False),
        ('paragraph', listing, 'List Paragraph', False),
    )
    content = ''
    for kind, sid, name, default in definitions:
        content += '<w:style w:type=' + quoteattr(kind) + ' w:styleId=' + quoteattr(sid)
        content += ' w:default="1"' if default else ''
        content += '><w:name w:val=' + quoteattr(name) + '/>'
        if kind == 'paragraph' and sid != normal:
            content += '<w:basedOn w:val=' + quoteattr(normal) + '/>'
        content += '</w:style>'
    return xml_document('styles', content)


def numbering_xml(num_id=7, abstract_id=3, nsid="10203040", template="50607080", bullet=False):
    """Return two-level numbering with deliberately controllable unstable ids."""
    levels = ''
    for level in (0, 1):
        fmt, text = ('bullet', '・') if bullet else ('decimal', '%' + str(level + 1) + '.')
        levels += ('<w:lvl w:ilvl="' + str(level) + '" w:tplc="A1B2C3D4" w:tentative="1">'
                   '<w:start w:val="1"/><w:numFmt w:val="' + fmt + '"/>'
                   '<w:lvlText w:val="' + text + '"/><w:lvlJc w:val="left"/></w:lvl>')
    return xml_document('numbering', '<w:abstractNum w:abstractNumId="' + str(abstract_id) + '">'
                        '<w:nsid w:val="' + nsid + '"/><w:tmpl w:val="' + template + '"/>'
                        + levels + '</w:abstractNum><w:num w:numId="' + str(num_id) + '">'
                        '<w:abstractNumId w:val="' + str(abstract_id) + '"/></w:num>')


def image(rel_id="rImage", doc_id=1, name="Synthetic picture", alt="試験用画像"):
    """Return a DrawingML inline picture referring to a caller-provided rel."""
    return ('<w:r><w:drawing><wp:inline><wp:extent cx="360000" cy="720000"/>'
            '<wp:docPr id="' + str(doc_id) + '" name=' + quoteattr(name) + ' descr=' + quoteattr(alt) + '/>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="Synthetic pixel"/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed=' + quoteattr(rel_id) + '/></pic:blipFill>'
            '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')


def hyperlink(text="試験リンク", rel_id="rLink", anchor=None):
    """Return a hyperlink with either an external rel or an internal anchor."""
    attr = 'w:anchor=' + quoteattr(anchor) if anchor is not None else 'r:id=' + quoteattr(rel_id)
    return '<w:hyperlink ' + attr + '>' + run(text) + '</w:hyperlink>'


def tracked(text, kind="ins", ident=5):
    """Return synthetic insertion or deletion markup."""
    tag = 'delText' if kind in ('del', 'moveFrom') else 't'
    return ('<w:' + kind + ' w:id="' + str(ident) + '" w:author="試験者" w:date="2026-01-02T03:04:05Z">'
            '<w:r><w:' + tag + '>' + escape(text) + '</w:' + tag + '></w:r></w:' + kind + '>')


def relationships(items):
    """Wrap relationship tuples, accepting suffixes or complete type URIs."""
    if isinstance(items, str):
        return items if '<Relationships' in items else '<Relationships xmlns="' + PR + '">' + items + '</Relationships>'
    content = ''
    for item in items:
        ident, kind, target = item[:3]
        kind = kind if '://' in kind else R + '/' + kind
        content += '<Relationship Id=' + quoteattr(ident) + ' Type=' + quoteattr(kind) + ' Target=' + quoteattr(target)
        if len(item) > 3:
            content += ' TargetMode=' + quoteattr(item[3])
        content += '/>'
    return '<Relationships xmlns="' + PR + '">' + content + '</Relationships>'


def zip_parts(parts, comment=b""):
    """Serialize named parts in insertion order with deterministic ZIP metadata."""
    output = io.BytesIO()
    with zipfile.ZipFile(output, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
        for name, value in parts.items():
            info = zipfile.ZipInfo(name, (2001, 2, 3, 4, 5, 6))
            info.compress_type = zipfile.ZIP_DEFLATED
            archive.writestr(info, value.encode('utf-8') if isinstance(value, str) else value)
        archive.comment = comment
    return output.getvalue()


def unpack(data):
    """Return the ZIP parts for deliberate corruption in refusal tests."""
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def make_docx(body='', *, styles=None, numbering=None, headers=None, footers=None,
              footnotes=None, endnotes=None, comments=None, rels=(), parts=None,
              japanese_ids=True, main_type='docx', main='word/document.xml',
              root_rels=None, content_types=None, signature=False, vba=False, comment=b''):
    """Build a synthetic Transitional Word ZIP from literal XML fragments.

    Header/footer keys are roles ('default', 'first', 'even') or (role, section).
    A value is a fragment, or (part_name, fragment) to vary physical part names.
    Supplied parts replace generated parts and may include raw bytes. A literal
    full styles/numbering/notes root is accepted in addition to inner XML.
    """
    package = {}
    main_kind = MAIN_TYPES.get(main_type, main_type)
    package_rels = [('rOffice', 'officeDocument', main)]
    document_rels = list(rels) if not isinstance(rels, str) else None
    generated_rels = []
    overrides = {main: main_kind}
    body = body or p()
    if not re.search(r'<w:sectPr(?:\s[^>]*)?>[\s\S]*?</w:sectPr>\s*$|<w:sectPr(?:\s[^>]*)?/>\s*$', body):
        body += '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr>'
    role_refs = {}
    for kind, collection in (('header', headers), ('footer', footers)):
        for index, (role_key, value) in enumerate((collection or {}).items(), 1):
            role, ordinal = role_key if isinstance(role_key, tuple) else (role_key, 1)
            part, fragment = value if isinstance(value, tuple) else ('word/' + kind + str(index) + '.xml', value)
            ident = 'r' + kind.title() + str(index)
            package[part] = fragment if fragment.lstrip().startswith('<?xml') else xml_document('hdr' if kind == 'header' else 'ftr', fragment)
            generated_rels.append((ident, kind, part.removeprefix('word/')))
            overrides[part] = 'application/vnd.openxmlformats-officedocument.wordprocessingml.' + kind + '+xml'
            role_refs.setdefault(ordinal, '')
            role_refs[ordinal] += '<w:' + kind + 'Reference w:type=' + quoteattr(role) + ' r:id=' + quoteattr(ident) + '/>'
    section_index = 0

    def add_references(match):
        nonlocal section_index
        section_index += 1
        fragment = match.group(0)
        refs = role_refs.get(section_index, '')
        if fragment.endswith('/>'):
            return fragment[:-2] + '>' + refs + '</w:sectPr>'
        return fragment.replace('>', '>' + refs, 1)

    body = re.sub(r'<w:sectPr(?:\s[^>]*)?/>|<w:sectPr(?:\s[^>]*)?>[\s\S]*?</w:sectPr>', add_references, body)
    package[main] = xml_document('document', '<w:body>' + body + '</w:body>')
    styles = default_styles(japanese_ids) if styles is None else styles
    if styles is not False:
        package['word/styles.xml'] = styles if '<w:styles' in styles else xml_document('styles', styles)
        generated_rels.append(('rStyles', 'styles', 'styles.xml'))
        overrides['word/styles.xml'] = 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml'
    if numbering is not None and numbering is not False:
        package['word/numbering.xml'] = numbering if '<w:numbering' in numbering else xml_document('numbering', numbering)
        generated_rels.append(('rNumbering', 'numbering', 'numbering.xml'))
        overrides['word/numbering.xml'] = 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml'
    for singular, fragment in (('footnote', footnotes), ('endnote', endnotes), ('comment', comments)):
        if fragment is None:
            continue
        plural = singular + 's'
        part = 'word/' + plural + '.xml'
        package[part] = fragment if '<w:' + plural in fragment else xml_document(plural, fragment)
        generated_rels.append(('r' + plural.title(), plural, plural + '.xml'))
        overrides[part] = 'application/vnd.openxmlformats-officedocument.wordprocessingml.' + plural + '+xml'
    if vba or main_type in ('docm', 'dotm'):
        package['word/vbaProject.bin'] = b'Synthetic VBA placeholder; never executable.'
        generated_rels.append(('rVba', 'http://schemas.microsoft.com/office/2006/relationships/vbaProject', 'vbaProject.bin'))
        overrides['word/vbaProject.bin'] = 'application/vnd.ms-office.vbaProject'
    if signature:
        package['_xmlsignatures/sig1.xml'] = '<Signature xmlns="http://www.w3.org/2000/09/xmldsig#"/>'
    package['_rels/.rels'] = relationships(package_rels) if root_rels is None else relationships(root_rels)
    rels_name = main.rsplit('/', 1)[0] + '/_rels/' + main.rsplit('/', 1)[-1] + '.rels'
    if document_rels is None:
        value = relationships(rels)
        package[rels_name] = value.replace('</Relationships>', relationships(generated_rels).split('>', 1)[1])
    else:
        package[rels_name] = relationships(document_rels + generated_rels)
    defaults = {'rels': 'application/vnd.openxmlformats-package.relationships+xml', 'xml': 'application/xml', 'png': 'image/png'}
    content = '<Types xmlns="' + CT + '">'
    content += ''.join('<Default Extension=' + quoteattr(ext) + ' ContentType=' + quoteattr(kind) + '/>' for ext, kind in defaults.items())
    content += ''.join('<Override PartName=' + quoteattr('/' + part) + ' ContentType=' + quoteattr(kind) + '/>' for part, kind in overrides.items())
    package['[Content_Types].xml'] = content_types if content_types is not None else content + '</Types>'
    package.update(parts or {})
    return zip_parts(package, comment)


def specification_docx(updated=False, japanese_ids=True):
    """Mimic a Word-authored specification's structure using invented content.

    The fixture includes a multi-paragraph TOC field, hidden targets, role-paired
    headers, footnotes, lists, a table, and an image. No real document content is
    used, and the alternate version has deliberate user-visible edits.
    """
    heading = '1' if japanese_ids else 'Heading1'
    listing = 'SyntheticJaList' if japanese_ids else 'ListParagraph'
    number = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr>'
    body = p('架空の試験計画', style=heading)
    body += p(runs='<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-2" \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>')
    entry = hyperlink('試験の目的', anchor='_TocSyntheticA')
    entry += '<w:r><w:tab/><w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve"> PAGEREF _TocSyntheticA \\h </w:instrText><w:fldChar w:fldCharType="separate"/><w:t>1</w:t><w:fldChar w:fldCharType="end"/></w:r>'
    body += p(runs=entry)
    if updated:
        body += p(runs=hyperlink('追加の条件', anchor='_TocSyntheticB'))
    body += p(runs='<w:r><w:fldChar w:fldCharType="end"/></w:r>')
    body += '<w:bookmarkStart w:id="41" w:name="_TocSyntheticA"/>' + p('試験の目的', style=heading) + '<w:bookmarkEnd w:id="41"/>'
    body += p('装置の温度を記録する。')
    body += p('注意事項', runs=run('注意事項', '<w:b/>' if updated else ''))
    body += p('準備する', style=listing, props=number) + p('計測する', style=listing, props=number)
    if updated:
        body += p('確認する', style=listing, props=number)
    body += table([['項目', '値'], ['電圧', '24' if updated else '12']])
    body += p(runs=run('測定結果') + '<w:r><w:footnoteReference w:id="8"/></w:r>')
    body += p(runs=image())
    if updated:
        body += '<w:bookmarkStart w:id="42" w:name="_TocSyntheticB"/>' + p('追加の条件', style=heading) + '<w:bookmarkEnd w:id="42"/>'
    notes = '<w:footnote w:id="-1" w:type="separator">' + p(runs='<w:r><w:separator/></w:r>') + '</w:footnote>'
    notes += '<w:footnote w:id="8">' + p('合成された脚注。') + '</w:footnote>'
    return make_docx(body, japanese_ids=japanese_ids, numbering=numbering_xml(),
                     headers={'first': p('試験計画・表紙'), 'default': p('試験計画・改訂' if updated else '試験計画・内部用')},
                     footers={'default': p(runs='<w:fldSimple w:instr="PAGE">' + run('1') + '</w:fldSimple>')},
                     footnotes=notes, rels=[('rImage', 'image', 'media/synthetic.png')],
                     parts={'word/media/synthetic.png': PNG})
