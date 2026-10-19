"""Read Word OOXML without rewriting package members.

Public records contain direct formatting only. Document.canon() and fmt_hash()
are also available to the staging planner; fresh_tree() always returns an
independent tree. All package paths use forward slashes.
"""

import copy
import hashlib
import io
import json
import posixpath
import re
import zipfile
from collections import namedtuple
from urllib.parse import unquote

from lxml import etree


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
XML = "http://www.w3.org/XML/1998/namespace"
EXTENSIONS = {".docx", ".docm", ".dotx", ".dotm"}
MAIN_TYPES = {
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
    "application/vnd.ms-word.document.macroEnabled.main+xml",
    "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
}
OLE_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_LIST_MARKERS = {
    "symbol": {0xB7: "•"},
    "wingdings": {
        0x6C: "●", 0x6E: "■", 0x6F: "□", 0x71: "❑", 0x75: "◆", 0x76: "❖",
        0xA7: "▪", 0xA8: "◻", 0xD8: "➢", 0xFC: "✔",
    },
}


class DocumentError(ValueError):
    """A package cannot be read or safely used for a partial edit."""


def _tag(name, ns=W):
    return "{" + ns + "}" + name


def _local(node):
    tag = node if isinstance(node, str) else node.tag
    return tag.rpartition("}")[2] if isinstance(tag, str) else ""


def _namespace(node):
    tag = node if isinstance(node, str) else node.tag
    return tag[1:tag.index("}")] if isinstance(tag, str) and tag.startswith("{") else ""


def _val(node, default=""):
    return node.get(_tag("val"), default) if node is not None else default


def _find(node, path):
    return node.find("/".join(_tag(p) for p in path.split("/"))) if node is not None else None


def _list_marker_text(level):
    """Render symbol-font markers for summaries without changing definitions."""
    text = _val(_find(level, "lvlText"))
    fonts = _find(level, "rPr/rFonts")
    font = fonts.get(_tag("ascii"), fonts.get(_tag("hAnsi"), "")) if fonts is not None else ""
    markers = _LIST_MARKERS.get(font.casefold(), {})
    return "".join(markers.get(ord(char) - 0xF000, "\\u%04x" % ord(char))
                   if "\uf000" <= char <= "\uf0ff" else char for char in text)


def _on(node):
    return node is not None and _val(node, "1").lower() not in ("0", "false", "off", "no")


def _unique(values):
    return tuple(dict.fromkeys(values))


def _json(value):
    return json.dumps(value, ensure_ascii=False, separators=(",", ":")).replace(
        "\u0085", "\\u0085").replace("\u2028", "\\u2028").replace("\u2029", "\\u2029")


def record_json(fields):
    """Encode a record on one physical line, including splitlines separators."""
    return _json(list(fields))


def _sha(value):
    return hashlib.sha256(value).hexdigest()


def _hash(value):
    return _sha(json.dumps(value, ensure_ascii=False, separators=(",", ":")).encode("utf-8"))[:8]


def _xml(data, part):
    try:
        tree = etree.parse(io.BytesIO(data), etree.XMLParser(
            resolve_entities=False, no_network=True, remove_blank_text=False))
    except (etree.XMLSyntaxError, ValueError) as error:
        raise DocumentError("Word の XML を読めません: " + part) from error
    if tree.docinfo.doctype:
        raise DocumentError("DTD を含む XML は扱えません: " + part)
    return tree


def _target(owner, target):
    target = unquote(target.split("#", 1)[0])
    path = posixpath.normpath(posixpath.join(posixpath.dirname(owner), target))
    if target.startswith("/"):
        path = posixpath.normpath(target.lstrip("/"))
    if path in ("", ".", "..") or path.startswith("../") or "\\" in path:
        raise DocumentError("不正な Word パーツの参照です")
    return path


def _rels_path(owner):
    return posixpath.join(posixpath.dirname(owner), "_rels", posixpath.basename(owner) + ".rels")


def _selected_children(node):
    """Read AlternateContent once, using the first Choice, never its Fallback."""
    for child in node:
        if child.tag == _tag("AlternateContent", MC):
            choice = child.find(_tag("Choice", MC))
            if choice is not None:
                yield from _selected_children(choice)
        elif isinstance(child.tag, str) and child.tag != _tag("Fallback", MC):
            yield child


def _walk(node):
    yield node
    for child in _selected_children(node):
        yield from _walk(child)


# Transitional Word content and property names supported by the read model.
# Unknown properties remain visible in the format digest; Microsoft property
# extensions and DrawingML extension payloads are preserved during copying.
_KNOWN_WORD = set("""
document body hdr ftr p pPr r rPr t tbl tblPr tblGrid gridCol tr trPr tc tcPr
sdt sdtPr sdtEndPr sdtContent customXml customXmlPr altChunk subDoc sectPr
pStyle rStyle tblStyle numPr ilvl numId numbering abstractNum num lvl lvlText
lvlJc numFmt start startOverride lvlOverride lvlRestart abstractNumId nsid tmpl
numStyleLink styleLink lvlPicBulletId numPicBullet suff tentative isLgl legacy
jc ind spacing contextualSpacing mirrorIndents adjustRightInd snapToGrid
pageBreakBefore keepNext keepLines widowControl suppressLineNumbers
pBdr shd tabs tab outlineLvl framePr textAlignment textboxTightWrap divId
bidi textDirection wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN
kinsoku suppressAutoHyphens suppressOverlap cnfStyle
b bCs i iCs u strike dstrike color highlight sz szCs rFonts vertAlign vanish
specVanish webHidden em bdr w position kern fitText caps smallCaps shadow
outline emboss imprint eastAsianLayout lang noProof rtl cs effect oMath
ins del moveFrom moveTo rPrChange pPrChange tblPrChange trPrChange tcPrChange
sectPrChange tblGridChange cellIns cellDel cellMerge
bookmarkStart bookmarkEnd permStart permEnd commentRangeStart commentRangeEnd
moveFromRangeStart moveFromRangeEnd moveToRangeStart moveToRangeEnd
customXmlInsRangeStart customXmlInsRangeEnd customXmlDelRangeStart customXmlDelRangeEnd
customXmlMoveFromRangeStart customXmlMoveFromRangeEnd customXmlMoveToRangeStart customXmlMoveToRangeEnd
proofErr lastRenderedPageBreak tab ptab br cr sym noBreakHyphen softHyphen
footnoteReference endnoteReference commentReference footnoteRef endnoteRef
annotationRef footnote endnote comment footnotes endnotes comments
fldChar instrText delInstrText delText fldSimple fldData hyperlink smartTag
smartTagPr attr bdo dir ruby rubyPr rubyBase rt rubyAlign hps hpsRaise hpsBase
lid drawing pict object control contentPart txbxContent pgNum
separator continuationSeparator continuationNotice dayLong dayShort monthLong
monthShort yearLong yearShort
alias tag id lock dataBinding temporary showingPlcHdr placeholder docPart
richText text comboBox dropDownList listItem date dateFormat lid storeMappedDataAs
calendar checked checkBox default size sizeAuto gallery docPartObj docPartList
citation equation picture group buildingBlock docPartGallery docPartUnique
headerReference footerReference printerSettings type pgSz pgMar paperSrc
pgBorders lnNumType pgNumType cols col formProt vAlign noEndnote titlePg
textDirection bidi rtlGutter docGrid footnotePr endnotePr footnotePos endnotePos
numStart numRestart numFmt pos
num numIdMacAtCleanup tblW tblInd tblBorders tblCellMar tblLook tblLayout
bidiVisual tblOverlap tblpPr tblCaption tblDescription tblStyleRowBandSize
tblStyleColBandSize tblPrEx tblPrExChange tblCellSpacing tblHeader cantSplit
trHeight gridBefore gridAfter wBefore wAfter hidden tblCellSpacing
cantSplitAfter tcW gridSpan vMerge hMerge tcBorders tcMar tcFitText noWrap
hideMark headers header tcPrChange tcW
firstRow lastRow firstColumn lastColumn noHBand noVBand
left right top bottom start end insideH insideV tl2br tr2bl between bar
before after inside outside
name style styles basedOn link next docDefaults latentStyles lsdException
qFormat semiHidden unhideWhenUsed uiPriority autoRedefine personal
personalCompose personalReply locked rsid default notTrueType panose1 charset
family pitch sig embedRegular embedBold embedItalic embedBoldItalic
commentReference annotationRef ffData statusText helpText enabled calcOnExit
entryMacro exitMacro textInput maxLength format formFieldData
numPrChange tblStylePr div divs marLeft marRight marTop marBottom
suppressTopSpacing suppressBottomSpacing
""".split())


Rel = namedtuple("Rel", "id type target external path")
Style = namedtuple("Style", "id type name default based_on link next num_id")
NumDef = namedtuple("NumDef", "abstract_id def_hash levels pic_bullet style_link")
NoteInfo = namedtuple("NoteInfo", "content_hash order text author date")


class ContentTypes:
    """Package content types; Override names and extensions are case-insensitive."""
    def __init__(self, root):
        self.defaults = {}
        self.overrides = {}
        for node in root:
            if _local(node) == "Default":
                self.defaults[node.get("Extension", "").lower()] = node.get("ContentType", "")
            elif _local(node) == "Override":
                self.overrides[node.get("PartName", "").lstrip("/").casefold()] = node.get("ContentType", "")

    def resolve(self, part):
        return self.overrides.get(part.lstrip("/").casefold(),
                                  self.defaults.get(part.rpartition(".")[2].lower()))


class Styles(dict):
    """Style ids mapped to metadata; by_key maps (type, casefold(name)) to ids."""
    def __init__(self, root=None, part=None):
        super().__init__()
        self.part = part
        self.by_key = {}
        self.nodes = {}
        if root is None:
            return
        for node in root.findall(_tag("style")):
            identifier = node.get(_tag("styleId"), "")
            kind = node.get(_tag("type"), "")
            name = _val(_find(node, "name"), identifier)
            style = Style(identifier, kind, name, node.get(_tag("default"), "0") in ("1", "true"),
                          _val(_find(node, "basedOn")), _val(_find(node, "link")),
                          _val(_find(node, "next")), _val(_find(node, "pPr/numPr/numId")))
            self[identifier] = style
            self.nodes[identifier] = node
            self.by_key.setdefault((kind, name.casefold()), []).append(identifier)


class Numbering(dict):
    """List ids mapped to normalized definitions, without computed item numbers."""
    def __init__(self):
        super().__init__()
        self.part = None
        self.nodes = {}
        self.abstracts = {}
        self.display_texts = {}


class Notes(dict):
    """(kind, XML id) mapped to content hashes and first-reference display order."""
    def __init__(self):
        super().__init__()
        self.nodes = {}
        self.parts = {}


class Line(namedtuple("LineBase", "kind label text format")):
    """An id-free immutable display record. key is its stable JSON encoding."""
    __slots__ = ()

    @property
    def key(self):
        return record_json(self)


class Refs(namedtuple("RefsBase", "styles num_ids rels bookmarks ref_names")):
    """Ordered unique references, retained even when normalization hides ids.

    styles holds (type, name), bookmarks holds names (including hidden names),
    and ref_names holds (kind, name). rels are ids local to the owning part.
    """
    __slots__ = ()

    @property
    def rel_ids(self):
        return self.rels


class RowShape(namedtuple("RowShapeBase", "line cells tr_index tc_indices")):
    """A table row and its raw child indices for lossless reconstruction."""
    __slots__ = ()


class TableShape(namedtuple("TableShapeBase", "signature table_line rows")):
    """Independent table/row/cell records; signature=None forbids slicing."""
    __slots__ = ()


class Unit:
    """An indivisible story range; section hosts address the original root.

    start/stop enclose attached markers and use an exclusive stop; indices
    and nodes give exact ownership when start/end markers interleave. A detached
    paragraph section has start==host and stop==host+1. nodes are original,
    read-only lxml elements; use Document.fresh_tree() before making edits.
    """
    def __init__(self, kind, start, stop, lines, refs, reasons=(), table=None, host=None, nodes=()):
        self.kind, self.start, self.stop, self.host = kind, start, stop, host
        self.lines = tuple(lines)
        self.key = "\n".join(line.key for line in self.lines)
        self.sim_text = "\n".join(line.text for line in self.lines)
        self.reasons = _unique(reasons)
        self.table, self.refs, self.nodes = table, refs, tuple(nodes)


class Story:
    """A body or referenced header/footer, identified by its first section role.

    roles preserves all (kind, type, section ordinal) references to a shared
    part; comparison can detect changes to roles without relying on filenames.
    """
    def __init__(self, identifier, title, part, root, roles=()):
        self.id, self.title, self.part, self.root = identifier, title, part, root
        self.roles = tuple(roles)
        self.units = ()


class Document:
    """Read-only package with original ZIP members and semantic records.

    bookmarks/ref_names cover selectable stories. all_bookmarks/all_ref_names
    also cover notes/comments, for whole-document reference validation.
    reference_roots is (id, Japanese title, part, root), with notes identified
    by footnotes/endnotes/comments; glossary is intentionally excluded.
    """
    def __init__(self, data, purpose="edit"):
        if purpose not in ("edit", "view"):
            raise DocumentError("Word 文書の読取目的が不正です")
        if data.startswith(OLE_SIGNATURE):
            raise DocumentError("暗号化された Word 文書（パスワード付き）または .doc 形式は扱えません")
        self.data, self.purpose = data, purpose
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as archive:
                self.infos = tuple(archive.infolist())
                if len({info.filename for info in self.infos}) != len(self.infos):
                    raise DocumentError("同名パーツが重複する Word 文書は扱えません")
                self.parts = {info.filename: archive.read(info) for info in self.infos}
                self.comment = archive.comment
        except DocumentError:
            raise
        except (OSError, RuntimeError, zipfile.BadZipFile, ValueError) as error:
            raise DocumentError("Word 文書を読めません: " + str(error)) from error
        # Parse before semantic checks, so DTD refusal takes precedence.
        self._trees = {part: _xml(value, part) for part, value in self.parts.items()
                       if part.endswith((".xml", ".rels")) or part == "[Content_Types].xml"}
        self._shas, self._rels, self._contexts, self._field_prefixes = {}, {}, {}, {}
        root = self._tree("_rels/.rels").getroot()
        office = [node for node in root if node.get("Type", "").endswith("/officeDocument") and node.get("TargetMode") != "External"]
        if len(office) != 1:
            raise DocumentError("Word の本文パーツの参照が不正です")
        if "purl.oclc.org" in office[0].get("Type", ""):
            raise DocumentError("Strict 形式の Word 文書は扱えません")
        self.main = _target("", office[0].get("Target", ""))
        main_root = self._tree(self.main).getroot()
        if (any("purl.oclc.org/ooxml" in n.get("Type", "") for p, tree in self._trees.items() if p.endswith(".rels") for n in tree.getroot())
                or "purl.oclc.org" in _namespace(main_root)
                or any(_local(key) == "conformance" and value.lower() == "strict"
                       for key, value in main_root.attrib.items())):
            raise DocumentError("Strict 形式の Word 文書は扱えません")
        self.content_types = ContentTypes(self._tree("[Content_Types].xml").getroot())
        main_type = self.content_types.resolve(self.main)
        if main_type not in MAIN_TYPES:
            raise DocumentError("Word 文書ではありません（本文パーツの種類: " + str(main_type) + "）")
        if main_root.tag != _tag("document") or _find(main_root, "body") is None:
            raise DocumentError("Word の本文 XML が不正です")
        if purpose == "edit" and (any(p.lower().startswith("_xmlsignatures/") for p in self.parts)
                or any("digital-signature" in n.get("Type", "") for p, t in self._trees.items()
                       if p.endswith(".rels") for n in t.getroot())):
            raise DocumentError("電子署名付きの Word 文書は段落単位で反映できません")
        # All rel files, including notes and orphaned headers, must be checked.
        for part in self._trees:
            if part.endswith(".rels"):
                identifiers = [n.get("Id", "") for n in self._trees[part].getroot()]
                if len(identifiers) != len(set(identifiers)):
                    owner = posixpath.join(posixpath.dirname(posixpath.dirname(part)), posixpath.basename(part)[:-5]) if part != "_rels/.rels" else "_rels/.rels"
                    raise DocumentError("relationship ID が重複しています: " + owner)
        self.styles = Styles()
        self.numbering, self.notes = Numbering(), Notes()
        style_part = self.related_part(self.main, "styles")
        if style_part:
            self.styles = Styles(self._tree(style_part).getroot(), style_part)
        self._load_numbering()
        self._load_notes()
        self.glossary_part = self.related_part(self.main, "glossaryDocument")
        self.glossary_names = set()
        if self.glossary_part:
            self.glossary_names = {_val(_find(n, "docPartPr/name")) for n in self._tree(self.glossary_part).getroot().iter(_tag("docPart"))}
        self.stories = self._stories(main_root)
        # Selectable stories retain their documented maps; the all_* maps
        # include notes for rebuild validation without inventing picker stories.
        roots = [(s.id, s.title, s.part, s.root) for s in self.stories]
        for kind, part in self.notes.parts.items():
            roots.append((kind + "s", {"footnote": "脚注", "endnote": "文末脚注", "comment": "コメント"}[kind],
                          part, self._tree(part).getroot()))
        self.reference_roots = tuple(roots)
        self.bookmarks, self.ref_names = {}, []
        self.all_bookmarks, self.all_ref_names = {}, []
        story_ids = {s.id for s in self.stories}
        for identifier, title, part, root in roots:
            seen_ids, seen_ends = set(), set()
            for node in _walk(root):
                if node.tag == _tag("bookmarkStart"):
                    bookmark_id, name = node.get(_tag("id"), ""), node.get(_tag("name"), "")
                    if purpose == "edit" and (bookmark_id in seen_ids or name.casefold() in self.all_bookmarks):
                        raise DocumentError("ブックマークの ID が重複しています: " + title)
                    seen_ids.add(bookmark_id)
                    self.all_bookmarks.setdefault(name.casefold(), []).append((identifier, bookmark_id))
                    if identifier in story_ids:
                        self.bookmarks.setdefault(name.casefold(), []).append((identifier, bookmark_id))
                elif node.tag == _tag("bookmarkEnd"):
                    bookmark_id = node.get(_tag("id"), "")
                    if purpose == "edit" and bookmark_id in seen_ends:
                        raise DocumentError("ブックマークの ID が重複しています: " + title)
                    seen_ends.add(bookmark_id)
            references = [(identifier, kind, name) for kind, name in self.references(root)]
            self.all_ref_names.extend(references)
            if identifier in story_ids:
                self.ref_names.extend(references)
        self._order_notes()
        for story in self.stories:
            story.units = self._units(story)

    def _tree(self, part):
        if part not in self.parts:
            raise DocumentError("Word のパーツがありません: " + part)
        if part not in self._trees:
            self._trees[part] = _xml(self.parts[part], part)
        return self._trees[part]

    def fresh_tree(self, part):
        """Parse a new independent tree, leaving the original package untouched."""
        if part not in self.parts:
            raise DocumentError("Word のパーツがありません: " + part)
        return _xml(self.parts[part], part)

    def part_sha(self, part):
        """Return the cached full SHA-256 digest of original member bytes."""
        if part not in self._shas:
            self._shas[part] = _sha(self.parts[part])
        return self._shas[part]

    def rels(self, part):
        """Return relationships owned by part (use an empty part for package rels)."""
        if part not in self._rels:
            path = _rels_path(part) if part else "_rels/.rels"
            values = {}
            if path in self.parts:
                for node in self._tree(path).getroot():
                    identifier, target = node.get("Id", ""), node.get("Target", "")
                    external = node.get("TargetMode") == "External"
                    values[identifier] = Rel(identifier, node.get("Type", ""), target,
                                              external, None if external else _target(part, target))
            self._rels[part] = values
        return self._rels[part]

    def related_part(self, part, suffix):
        """Find the first internal relationship of the requested type suffix."""
        return next((rel.path for rel in self.rels(part).values()
                     if not rel.external and rel.type.endswith("/" + suffix)), None)

    def _stories(self, main_root):
        body = Story("body", "本文", self.main, _find(main_root, "body"))
        descriptors = {}
        for ordinal, section in enumerate(body.root.iter(_tag("sectPr")), 1):
            for node in section:
                if _local(node) not in ("headerReference", "footerReference"):
                    continue
                kind = "header" if _local(node) == "headerReference" else "footer"
                role = node.get(_tag("type"), "default")
                rel = self.rels(self.main).get(node.get(_tag("id", R), ""))
                if rel is None or rel.external or rel.path not in self.parts:
                    raise DocumentError("Word のヘッダー・フッターの参照が不正です")
                descriptors.setdefault(rel.path, []).append((kind, role, ordinal))
        result = [body]
        titles = {"header": "ヘッダー", "footer": "フッター"}
        roles = {"default": "既定", "first": "先頭ページ", "even": "偶数ページ"}
        for part, references in descriptors.items():
            kind, role, ordinal = references[0]
            title = titles[kind] + "（セクション " + str(ordinal) + "・" + roles.get(role, role) + "）"
            result.append(Story(kind + ":" + str(ordinal) + ":" + role, title, part,
                                self._tree(part).getroot(), references))
        return tuple(result)

    def _load_numbering(self):
        part = self.related_part(self.main, "numbering")
        self.numbering.part = part
        if not part:
            return
        root = self._tree(part).getroot()
        abstracts = {_val(n, n.get(_tag("abstractNumId"), "")): n
                     for n in root.findall(_tag("abstractNum"))}
        self.numbering.abstracts = abstracts
        for node in root.findall(_tag("num")):
            identifier = node.get(_tag("numId"), "")
            abstract_id = _val(_find(node, "abstractNumId"))
            abstract = abstracts.get(abstract_id)
            self.numbering.nodes[identifier] = node
            if abstract is None:
                self.numbering[identifier] = NumDef(abstract_id, "missing", {}, False, False)
                continue
            levels = {n.get(_tag("ilvl"), "0"): (_val(_find(n, "numFmt")), _val(_find(n, "lvlText")))
                      for n in abstract.findall(_tag("lvl"))}
            display_texts = {n.get(_tag("ilvl"), "0"): _list_marker_text(n)
                             for n in abstract.findall(_tag("lvl"))}
            for override in node.findall(_tag("lvlOverride")):
                level = _find(override, "lvl")
                if level is not None:
                    levels[override.get(_tag("ilvl"), "0")] = (
                        _val(_find(level, "numFmt")), _val(_find(level, "lvlText")))
                    display_texts[override.get(_tag("ilvl"), "0")] = _list_marker_text(level)
            self.numbering.display_texts[identifier] = display_texts
            definitions = [copy.deepcopy(abstract), copy.deepcopy(node)]
            for definition in definitions:
                for element in list(definition.iter()):
                    if _local(element) in ("nsid", "tmpl", "abstractNumId") and element is not definition:
                        element.getparent().remove(element)
                        continue
                    for key in list(element.attrib):
                        if _local(key) in ("abstractNumId", "numId", "tplc", "tentative"):
                            del element.attrib[key]
            digest = _sha(_json([self.canon(n, part) for n in definitions]).encode("utf-8"))
            pic = any(_local(n) == "lvlPicBulletId" for root in (abstract, node) for n in root.iter())
            linked = any(_local(n) in ("numStyleLink", "styleLink") for root in (abstract, node) for n in root.iter())
            self.numbering[identifier] = NumDef(abstract_id, digest, levels, pic, linked)

    def _load_notes(self):
        for kind, suffix in (("footnote", "footnotes"), ("endnote", "endnotes"), ("comment", "comments")):
            part = self.related_part(self.main, suffix)
            if not part:
                continue
            self.notes.parts[kind] = part
            for node in self._tree(part).getroot():
                if _local(node) != kind:
                    continue
                identifier = node.get(_tag("id"), "")
                if kind != "comment" and (identifier in ("-1", "0") or node.get(_tag("type")) in ("separator", "continuationSeparator")):
                    continue
                key = kind, identifier
                self.notes.nodes[key] = node
                clone = copy.deepcopy(node)
                for attr in ("id",):
                    clone.attrib.pop(_tag(attr), None)
                digest = _sha(_json(self.canon(clone, part)).encode("utf-8") + self.display_text(node, part).encode("utf-8"))
                self.notes[key] = NoteInfo(digest, 0, self._block_text(node, part),
                                           node.get(_tag("author"), ""), node.get(_tag("date"), ""))

    def _order_notes(self):
        counts = {"footnote": 0, "endnote": 0, "comment": 0}
        roots = [s.root for s in self.stories] + list(self.notes.nodes.values())
        keys = []
        for root in roots:
            for node in _walk(root):
                kind = {"footnoteReference": "footnote", "endnoteReference": "endnote",
                        "commentReference": "comment", "commentRangeStart": "comment"}.get(_local(node))
                key = (kind, node.get(_tag("id"), ""))
                if kind and key in self.notes and key not in keys:
                    keys.append(key)
        keys.extend(key for key in self.notes if key not in keys)
        for key in keys:
            counts[key[0]] += 1
            self.notes[key] = self.notes[key]._replace(order=counts[key[0]])

    def _context(self, node):
        root = node.getroottree().getroot()
        if root in self._contexts:
            return self._contexts[root]
        starts = {n.get(_tag("id"), ""): n.get(_tag("name"), "")
                  for n in _walk(root) if n.tag == _tag("bookmarkStart")}
        perms = {n.get(_tag("id"), ""): n.get(_tag("edGrp"), n.get(_tag("ed"), ""))
                 for n in _walk(root) if n.tag == _tag("permStart")}
        self._contexts[root] = (starts, perms)
        return starts, perms


    def _field_state(self, node):
        """Return field-result state entering a fragment, including prior blocks."""
        root = node.getroottree().getroot()
        if root not in self._field_prefixes:
            states, stack = {}, []
            def visit(element):
                states[element] = tuple(stack)
                if element.tag in (_tag("del"), _tag("moveFrom"), _tag("rt")):
                    return
                if element.tag == _tag("fldChar"):
                    mode = element.get(_tag("fldCharType"))
                    if mode == "begin":
                        stack.append(False)
                    elif mode == "separate" and stack:
                        stack[-1] = True
                    elif mode == "end" and stack:
                        stack.pop()
                for child in _selected_children(element):
                    visit(child)
            visit(root)
            self._field_prefixes[root] = states
        return self._field_prefixes[root].get(node, ())

    def normalized(self, node, detach_sections=True):
        """Return an independent tree fragment after N1-N3 structural cleanup.

        Relationship/style/note tokens are applied by canon(), not written into
        this XML. The original ids remain available for the staging planner.
        """
        starts, _ = self._context(node)
        clone = copy.deepcopy(node)
        self._field_prefixes[clone] = {clone: self._field_state(node)}
        drop = {"proofErr", "lastRenderedPageBreak", "noProof", "lang", "cnfStyle", "printerSettings"}

        def clean(element):
            for child in list(element):
                if not isinstance(child.tag, str):
                    element.remove(child)
                    continue
                local, ns = _local(child), _namespace(child)
                hidden = (ns == W and ((local == "bookmarkStart" and child.get(_tag("name"), "").startswith("_"))
                          or (local == "bookmarkEnd" and starts.get(child.get(_tag("id"), ""), "").startswith("_"))))
                creation = (ns == A and local == "ext" and child.get("uri", "").upper()
                            == "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}")
                if ((ns == W and local in drop) or hidden or creation
                        or (detach_sections and child.tag == _tag("sectPr") and element.tag == _tag("pPr"))):
                    element.remove(child)
                    continue
                if child.tag == _tag("AlternateContent", MC):
                    choice = child.find(_tag("Choice", MC))
                    position = element.index(child)
                    if choice is not None:
                        clean(choice)
                        for inner in list(choice):
                            element.insert(position, inner)
                            position += 1
                    element.remove(child)
                    continue
                if child.tag == _tag("Fallback", MC):
                    element.remove(child)
                    continue
                clean(child)
                if child.tag == _tag("rFonts") and set(child.attrib).issubset({_tag("hint")}):
                    element.remove(child)
                elif child.tag in (_tag("rPr"), _tag("pPr"), _tag("extLst", A)) and not len(child) and not child.attrib:
                    element.remove(child)
                elif child.tag == _tag("r") and not any(c.tag != _tag("rPr") for c in child):
                    element.remove(child)
            if element.tag in (_tag("t"), _tag("delText")):
                if element.get(_tag("space", XML)) != "preserve":
                    element.text = (element.text or "").strip(" \t\r\n")
            for key in list(element.attrib):
                local, ns = _local(key), _namespace(key)
                if ((ns == W and local.startswith("rsid"))
                        or (ns.endswith("/wordml") and local in ("paraId", "textId"))
                        or (ns.endswith("/wordprocessingDrawing") and local in ("anchorId", "editId"))
                        or ("/word/" in ns and local == "dateUtc") or key == _tag("space", XML)
                        or (local in ("id", "name") and _local(element) in ("docPr", "cNvPr")
                            and _namespace(element) != W)):
                    del element.attrib[key]
            if element.tag == _tag("tblLook"):
                try:
                    bits = int(_val(element, "0"), 16)
                except ValueError:
                    bits = 0
                flags = ("firstRow", "lastRow", "firstColumn", "lastColumn", "noHBand", "noVBand")
                for flag, mask in zip(flags, (0x20, 0x40, 0x80, 0x100, 0x200, 0x400)):
                    explicit = element.get(_tag(flag))
                    element.set(_tag(flag), "1" if (explicit.lower() not in ("0", "false", "off")
                                if explicit is not None else bits & mask) else "0")
                element.attrib.pop(_tag("val"), None)
            # Merge only siblings. A formatting boundary must survive even if
            # its affected text happens to have the same spelling.
            previous = None
            for child in list(element):
                if previous is not None and child.tag == previous.tag == _tag("r"):
                    before, after = _find(previous, "rPr"), _find(child, "rPr")
                    same = _structure(before) == _structure(after)
                    if same:
                        for inner in list(child):
                            if inner.tag != _tag("rPr"):
                                previous.append(inner)
                        element.remove(child)
                        merge_text(previous)
                        continue
                previous = child
            merge_text(element)

        def merge_text(element):
            previous = None
            for child in list(element):
                if (previous is not None and child.tag == previous.tag
                        and child.tag in (_tag("t"), _tag("instrText"), _tag("delText"), _tag("delInstrText"))):
                    previous.text = (previous.text or "") + (child.text or "")
                    element.remove(child)
                elif child.tag == _tag("t") and not (child.text or ""):
                    element.remove(child)
                else:
                    previous = child
        clean(clone)
        return clone

    def canon(self, node, part=None, detach_sections=True):
        """Return N1-N5 nested lists, with display-text lengths and semantic ids.

        part is the owner of local relationship ids (defaults to main). This
        does not expand style definitions or compute list item numbers.
        """
        part = part or self.main
        starts, perms = self._context(node)
        normalized = self.normalized(node, detach_sections)
        fields = list(self._field_state(node))

        def build(element, visible=True):
            local, ns = _local(element), _namespace(element)
            visible = visible and not (ns == W and local in ("del", "moveFrom", "rt"))
            if element.tag == _tag("fldChar"):
                mode = element.get(_tag("fldCharType"))
                if mode == "begin":
                    fields.append(False)
                elif mode == "separate" and fields:
                    fields[-1] = True
                elif mode == "end" and fields:
                    fields.pop()
            attrs = []
            for key, value in sorted(element.attrib.items()):
                name = _local(key)
                if key == _tag("id"):
                    if local in ("headerReference", "footerReference"):
                        pass
                    elif local in ("footnoteReference", "endnoteReference", "commentReference", "commentRangeStart", "commentRangeEnd"):
                        kind = "comment" if local.startswith("comment") else local.replace("Reference", "")
                        note = self.notes.get((kind, value))
                        value = ["#note", kind, note.content_hash if note else "missing"]
                    elif local == "bookmarkEnd":
                        value = starts.get(value, "missing")
                    elif local == "permEnd":
                        value = perms.get(value, "missing")
                    elif (local in ("bookmarkStart", "permStart", "ins", "del", "moveFrom", "moveTo", "cellIns", "cellDel", "cellMerge")
                          or local.endswith("PrChange") or local.startswith(("moveFromRange", "moveToRange", "customXml"))):
                        continue
                if element.tag == _tag("id") and element.getparent() is not None and element.getparent().tag == _tag("sdtPr") and key == _tag("val"):
                    continue
                if key == _tag("val") and local in ("pStyle", "rStyle", "tblStyle"):
                    style = self.styles.get(value)
                    value = ["#style", style.type, style.name] if style else ["#style", "missing", value]
                elif key == _tag("val") and local == "numId":
                    definition = self.numbering.get(value)
                    value = ["#num", "none" if value == "0" else definition.def_hash if definition else "missing"]
                elif _namespace(key) == R:
                    if local in ("headerReference", "footerReference"):
                        value = ["#hdr", local.replace("Reference", ""), element.get(_tag("type"), "default")]
                    else:
                        rel = self.rels(part).get(value)
                        if rel is None:
                            value = ["#rel", "missing"]
                        elif rel.external:
                            value = ["#rel", rel.type.rsplit("/", 1)[-1], "ext", rel.target]
                        else:
                            value = ["#rel", rel.type.rsplit("/", 1)[-1], "int",
                                     self.part_sha(rel.path) if rel.path in self.parts else "missing",
                                     self.content_types.resolve(rel.path)]
                attrs.append([key, value])
            result = [element.tag, attrs]
            if element.text and (local in ("t", "delText", "instrText", "delInstrText") or element.text.strip()):
                displayed = visible and all(fields) and local == "t" and ns in (W, M)
                result.append(["#t", len(element.text)] if displayed else ["#v", element.text])
            result.extend(build(child, visible) for child in element if isinstance(child.tag, str))
            return result
        return build(normalized)

    def fmt_hash(self, node, part=None, detach_sections=True):
        """Return the eight-hex format digest used in public Line records."""
        return _hash(self.canon(node, part, detach_sections))

    def _style_name(self, node, kind="pStyle"):
        identifier = _val(_find(node, "pPr/" + kind)) if kind == "pStyle" else _val(_find(node, "tblPr/" + kind))
        style = self.styles.get(identifier)
        return style.name if style else identifier

    def display_text(self, node, part=None, numbered_notes=False):
        """Render No-Markup text; hidden runs stay visible, deleted text does not."""
        part = part or self.main
        field_stack = list(self._field_state(node))

        def display(element):
            local, ns = _local(element), _namespace(element)
            if ns == W and local in ("del", "moveFrom", "delText", "delInstrText", "instrText", "rt", "txbxContent"):
                return ""
            if ns == W and local == "fldChar":
                mode = element.get(_tag("fldCharType"))
                if mode == "begin":
                    field_stack.append(False)
                elif mode == "separate" and field_stack:
                    field_stack[-1] = True
                elif mode == "end" and field_stack:
                    field_stack.pop()
                return ""
            if local == "t" and ns in (W, M):
                if field_stack and not all(field_stack):
                    return ""
                text = element.text or ""
                return text if element.get(_tag("space", XML)) == "preserve" or ns == M else text.strip(" \t\r\n")
            if ns == W:
                if local in ("tab", "ptab"):
                    return "\t"
                if local in ("br", "cr"):
                    return {"page": "［改ページ］", "column": "［段区切り］"}.get(element.get(_tag("type")), "\n")
                if local in ("noBreakHyphen", "softHyphen"):
                    return "\u2011" if local == "noBreakHyphen" else "\u00ad"
                if local == "sym":
                    return "［記号 " + element.get(_tag("font"), "") + " " + element.get(_tag("char"), "") + "］"
                if local == "pgNum":
                    return "［ページ番号］"
                if local in ("footnoteReference", "endnoteReference", "commentReference"):
                    kind = local.replace("Reference", "")
                    note = self.notes.get((kind, element.get(_tag("id"), "")))
                    label = {"footnote": "脚注", "endnote": "文末脚注", "comment": "コメント"}[kind]
                    return "［" + label + (" " + str(note.order) if numbered_notes and note and note.order else "") + "］"
                if local == "tbl":
                    return self._nested_table_text(element, part, numbered_notes)
                if local == "object":
                    return "［オブジェクト］"
                if local == "pict":
                    return "［図形］"
            if ns == M and local in ("oMath", "oMathPara"):
                return "［数式 " + "".join(n.text or "" for n in element.iter(_tag("t", M))) + "］"
            if local == "symEx" and "word" in ns:
                value = element.get(_tag("char", ns), element.get("char", ""))
                try:
                    return chr(int(value, 16))
                except (ValueError, OverflowError):
                    return "［記号 " + value + "］"
            if local == "drawing" and ns == W:
                images = self._images(element, part)
                if images:
                    return "".join("［画像 " + digest + "］" for digest, _, _, _ in images)
                nodes = list(_walk(element))
                if any(_local(n) == "chart" for n in nodes):
                    return "［グラフ］"
                return "［図形］"
            return "".join(display(child) for child in _selected_children(element))
        return display(node)

    def _images(self, node, part):
        images = []
        for element in _walk(node):
            if element.tag != _tag("blip", A):
                continue
            rel = self.rels(part).get(element.get(_tag("embed", R), element.get(_tag("link", R), "")))
            if rel is None:
                digest = "missing"
            else:
                digest = _sha(rel.target.encode("utf-8"))[:8] if rel.external else self.part_sha(rel.path)[:8] if rel.path in self.parts else "missing"
            drawing = next((n for n in element.iterancestors() if n.tag == _tag("drawing")), node)
            extent = next((n for n in _walk(drawing) if n.tag == _tag("extent", WP)), None)
            docpr = next((n for n in _walk(drawing) if n.tag == _tag("docPr", WP)), None)
            width = _decimal(extent.get("cx", "0"), 360000) if extent is not None else "0"
            height = _decimal(extent.get("cy", "0"), 360000) if extent is not None else "0"
            alt = docpr.get("descr", docpr.get("title", "")) if docpr is not None else ""
            images.append((digest, width, height, alt))
        return images

    def _block_text(self, node, part, numbered_notes=False):
        """Join block text, including blocks inside controls and revision wrappers."""
        blocks = []
        for child in _selected_children(node):
            if child.tag in (_tag("p"), _tag("tbl")):
                blocks.append(self.display_text(child, part, numbered_notes))
            elif child.tag == _tag("sdt"):
                content = _find(child, "sdtContent")
                if content is not None:
                    blocks.append(self._block_text(content, part, numbered_notes))
            elif child.tag in (_tag("customXml"), _tag("ins"), _tag("moveTo")):
                blocks.append(self._block_text(child, part, numbered_notes))
        return "\n".join(blocks) if blocks else self.display_text(node, part, numbered_notes)

    def _nested_table_text(self, table, part, numbered_notes=False):
        rows = []
        for row in table.findall(_tag("tr")):
            rows.append(" | ".join(self._block_text(cell, part, numbered_notes)
                                   for cell in row.findall(_tag("tc"))))
        return "［表: " + " / ".join(rows) + "］"

    def references(self, node):
        """Return ordered (TOC/相互参照/文書内リンク, bookmark name) pairs."""
        references, stack = [], []

        def fields(code):
            words = re.findall(r'"[^"\\]*(?:\\.[^"\\]*)*"|[^\s]+', code.strip())
            if not words:
                return
            command = words[0].upper()
            if command in ("REF", "PAGEREF", "NOTEREF", "GOTOBUTTON") and len(words) > 1 and not words[1].startswith("\\"):
                name = words[1].strip('"')
                references.append(("目次" if command == "PAGEREF" and name.casefold().startswith("_toc") else "相互参照", name))
            elif command == "HYPERLINK":
                for index, word in enumerate(words[:-1]):
                    if word.lower() == "\\l":
                        references.append(("文書内リンク", words[index + 1].strip('"')))
        for element in (element for root in (node if isinstance(node, (list, tuple)) else (node,)) for element in _walk(root)):
            if element.tag == _tag("hyperlink") and element.get(_tag("anchor")):
                name = element.get(_tag("anchor"))
                references.append(("目次" if name.casefold().startswith("_toc") else "文書内リンク", name))
            elif element.tag == _tag("fldSimple"):
                fields(element.get(_tag("instr"), ""))
            elif element.tag == _tag("fldChar"):
                mode = element.get(_tag("fldCharType"))
                if mode == "begin":
                    stack.append([])
                elif mode in ("separate", "end") and stack:
                    fields("".join(stack[-1]))
                    stack[-1] = []
                    if mode == "end":
                        stack.pop()
            elif element.tag in (_tag("instrText"), _tag("delInstrText")) and stack:
                stack[-1].append(element.text or "")
        for code in stack:
            fields("".join(code))
        return _unique(references)

    def refs(self, nodes, part):
        """Collect source references for a unit or a table piece, in XML order."""
        nodes = tuple(nodes)
        styles, nums, rels, bookmarks = [], [], [], []
        names = self.references(nodes)
        for node in nodes:
            for element in _walk(node):
                local = _local(element)
                if local in ("pStyle", "rStyle", "tblStyle") and _namespace(element) == W:
                    identifier = _val(element)
                    style = self.styles.get(identifier)
                    styles.append((style.type, style.name) if style else ("missing", identifier))
                elif element.tag == _tag("numId"):
                    nums.append(_val(element))
                elif element.tag == _tag("bookmarkStart"):
                    bookmarks.append(element.get(_tag("name"), ""))
                for key, value in element.attrib.items():
                    if _namespace(key) == R:
                        rels.append(value)
        return Refs(_unique(styles), _unique(nums), _unique(rels), _unique(bookmarks), _unique(names))

    def _plain(self, node):
        if node.tag != _tag("p") or node.attrib or not all(self._field_state(node)):
            return False
        for child in node:
            if child.tag == _tag("pPr"):
                if child.attrib or any(n.tag != _tag("pStyle") for n in child):
                    return False
            elif child.tag == _tag("r"):
                if child.attrib or any(n.tag != _tag("t") or n.attrib for n in child):
                    return False
            else:
                return False
        return len(node.findall(_tag("r"))) <= 1

    def _format(self, node, part, plain=False, extra=(), detach_sections=True):
        normalized = self.normalized(node, detach_sections)
        tokens = list(extra) + self._tokens(normalized, part)
        if not (plain and self._plain(normalized)):
            tokens.append("#" + self.fmt_hash(node, part, detach_sections))
        return " ".join(_unique(tokens))

    def _tokens(self, node, part):
        tokens = []
        local = _local(node)
        if local == "tc":
            pr = _find(node, "tcPr")
            if _find(pr, "shd") is not None:
                tokens.append("網かけ")
            vertical = _find(pr, "vMerge")
            if vertical is not None:
                tokens.append("縦結合(開始)" if _val(vertical) == "restart" else "縦結合(続き)")
            span = _find(pr, "gridSpan")
            if span is not None:
                tokens.append("横結合" + _val(span) + "列")
            if _find(pr, "vAlign") is not None:
                tokens.append("セル内の配置")
            for paragraph in _walk(node):
                if paragraph.tag != _tag("p") or next((a for a in paragraph.iterancestors() if a.tag == _tag("tc")), None) is not node:
                    continue
                style = self._style_name(paragraph)
                if style:
                    tokens.append("スタイル「" + style + "」")
                tokens.extend(self._tokens(paragraph, part))
            return list(_unique(tokens))
        properties = _find(node, "pPr")
        if properties is not None:
            numpr = _find(properties, "numPr")
            if numpr is not None:
                identifier, level = _val(_find(numpr, "numId")), _val(_find(numpr, "ilvl"), "0")
                definition = self.numbering.get(identifier)
                fmt, text = definition.levels.get(level, ("", "")) if definition else ("", "")
                if identifier == "0":
                    tokens.append("番号なし")
                elif identifier:
                    text = self.numbering.display_texts.get(identifier, {}).get(level, text)
                    tokens.append(("箇条書き" if fmt == "bullet" else "番号") + "「" + text + "」L" + _increment(level))
                else:
                    tokens.append("リスト L" + _increment(level))
            alignment = _val(_find(properties, "jc"))
            if alignment:
                tokens.append({"left": "左揃え", "center": "中央揃え", "right": "右揃え", "both": "両端揃え",
                               "distribute": "均等割り付け", "start": "左揃え", "end": "右揃え"}.get(alignment, "配置(" + alignment + ")"))
            if _find(properties, "ind") is not None:
                tokens.append("インデント")
            spacing = _find(properties, "spacing")
            if spacing is not None:
                for attributes, label in ((("line", "lineRule"), "行間"), (("before", "beforeLines", "beforeAutospacing"), "段落前"),
                                           (("after", "afterLines", "afterAutospacing"), "段落後")):
                    if any(_tag(a) in spacing.attrib for a in attributes):
                        tokens.append(label)
            for name, label in (("pageBreakBefore", "段落前で改ページ"), ("keepNext", "次の段落と分離しない"),
                                ("keepLines", "段落を分割しない")):
                if _on(_find(properties, name)):
                    tokens.append(label)
            for name, label in (("pBdr", "段落罫線"), ("shd", "段落の網かけ"), ("tabs", "タブ位置")):
                if _find(properties, name) is not None:
                    tokens.append(label)
            outline = _find(properties, "outlineLvl")
            if outline is not None:
                tokens.append("アウトライン L" + _increment(_val(outline)))
            if _find(properties, "framePr") is not None:
                tokens.append("枠")
            for name, label in (("rPr/ins", "段落記号の挿入"), ("rPr/del", "段落記号の削除")):
                if _find(properties, name) is not None:
                    tokens.append(label)
            if any(_local(n).endswith("PrChange") for n in properties.iter()):
                tokens.append("書式の変更履歴")
        all_text = self.display_text(node, part)
        run_properties = []
        # Use normalized runs so splitting a run never changes the summary.
        for run in _walk(node):
            if run.tag != _tag("r"):
                continue
            rpr = _find(run, "rPr")
            if rpr is not None:
                for order, label in self._run_tokens(rpr):
                    run_properties.append((order, label, self.display_text(run, part)))
        labels = {}
        for order, label, text in run_properties:
            labels.setdefault((order, label), []).append(text)
        for (_, label), texts in sorted(labels.items(), key=lambda item: item[0][0]):
            if "".join(texts) == all_text:
                tokens.append(label)
            elif texts:
                tokens.append(label + "「" + _excerpt(texts[0]) + "」" + (
                    "ほか" + str(len(texts) - 1) + " か所" if len(texts) > 1 else ""))
        for element in _walk(node):
            name = _local(element)
            if name == "ruby" and _namespace(element) == W:
                base, ruby = _find(element, "rubyBase"), _find(element, "rt")
                text = self.display_text(base, part) if base is not None else ""
                reading = "".join(n.text or "" for n in ruby.iter(_tag("t"))) if ruby is not None else ""
                tokens.append("ルビ「" + text + "(" + reading + ")」")
        for digest, width, height, alt in self._images(node, part):
            tokens.append("画像「" + digest + "」" + width + "×" + height + "cm「" + alt + "」")
        codes, stack = [], []
        for element in _walk(node):
            name, ns = _local(element), _namespace(element)
            if ns != W:
                continue
            if name == "hyperlink":
                rel = self.rels(part).get(element.get(_tag("id", R), ""))
                target = rel.target if rel else "#" + element.get(_tag("anchor"), "")
                tokens.append("リンク「" + self.display_text(element, part) + "」→" + target)
            elif name == "fldSimple":
                codes.append(element.get(_tag("instr"), ""))
            elif name == "fldChar":
                mode = element.get(_tag("fldCharType"))
                if mode == "begin":
                    stack.append([])
                elif mode in ("separate", "end") and stack:
                    code = "".join(stack[-1]).strip()
                    if code:
                        codes.append(code)
                    stack[-1] = []
                    if mode == "end":
                        stack.pop()
            elif name == "instrText" and stack:
                stack[-1].append(element.text or "")
            elif name == "bookmarkStart" and not element.get(_tag("name"), "").startswith("_"):
                tokens.append("ブックマーク「" + element.get(_tag("name"), "") + "」")
            elif name == "sdt":
                tag = _val(_find(element, "sdtPr/tag"), _val(_find(element, "sdtPr/alias")))
                tokens.append("コンテンツ コントロール「" + tag + "」")
            elif name in ("ins", "del", "moveFrom", "moveTo"):
                text = ("".join(n.text or "" for n in element.iter(_tag("delText")))
                        if name in ("del", "moveFrom") else self.display_text(element, part))
                tokens.append("変更履歴(" + ("削除" if name in ("del", "moveFrom") else "挿入") + "「" + _excerpt(text) + "」)")
        for code in codes:
            words = code.strip().split()
            if words:
                tokens.append("フィールド「" + words[0] + "」")
        if local in ("tbl", "table-geometry"):
            pr = _find(node, "tblPr")
            for name, label in (("tblW", "幅"), ("tblBorders", "罫線"), ("jc", "配置")):
                if _find(pr, name) is not None:
                    tokens.append(label)
        if local in ("tr", "row-properties"):
            pr = _find(node, "trPr")
            for name, label in (("trHeight", "行の高さ"), ("tblHeader", "見出し行の繰り返し"), ("cantSplit", "行を分割しない")):
                element = _find(pr, name)
                if element is not None and (name == "trHeight" or _on(element)):
                    tokens.append(label)
        if local == "sectPr":
            margins = _find(node, "pgMar")
            if margins is not None:
                tokens.append("余白 " + " ".join(label + _decimal(margins.get(_tag(name), "0"), 1440 / 25.4)
                                                 for name, label in (("top", "上"), ("bottom", "下"), ("left", "左"), ("right", "右"))) + " mm")
            columns = _find(node, "cols")
            if columns is not None:
                tokens.append("段組み" + columns.get(_tag("num"), str(max(1, len(columns)))))
            if _find(node, "lnNumType") is not None:
                tokens.append("行番号")
            page = _find(node, "pgNumType")
            if page is not None and page.get(_tag("start")):
                tokens.append("ページ番号(開始" + page.get(_tag("start")) + ")")
            if _on(_find(node, "titlePg")):
                tokens.append("先頭ページ別指定")
            direction = _val(_find(node, "textDirection"))
            if direction and direction not in ("lrTb", "lr"):
                tokens.append("縦書き")
        return list(_unique(tokens))

    def _run_tokens(self, properties):
        labels = []
        order = 0
        def add(label):
            nonlocal order
            labels.append((order, label))
        for name, label in (("b", "太字"), ("i", "斜体"), ("u", "下線"), ("strike", "取り消し線"), ("dstrike", "二重取り消し線")):
            element = _find(properties, name)
            if element is not None:
                value = _val(element, "single" if name == "u" else "1")
                if value in ("0", "false", "off", "none"):
                    if name in ("b", "i", "u"):
                        add(label + "なし")
                elif name == "u" and value != "single":
                    add(label + "(" + value + ")")
                else:
                    add(label)
            order += 1
        color = _find(properties, "color")
        if color is not None:
            theme = color.get(_tag("themeColor"))
            add("色(テーマ:" + theme + ")" if theme else "色(自動)" if _val(color) == "auto" else "色#" + _val(color).upper())
        order += 1
        highlight = _find(properties, "highlight")
        if highlight is not None:
            add("蛍光ペン(" + _val(highlight) + ")")
        order += 1
        shading = _find(properties, "shd")
        if shading is not None:
            add("網かけ#" + shading.get(_tag("fill"), "auto").upper())
        order += 1
        size = _find(properties, "sz")
        if size is None:
            size = _find(properties, "szCs")
        if size is not None:
            add("サイズ" + _decimal(_val(size), 2) + "pt")
        order += 1
        fonts = _find(properties, "rFonts")
        if fonts is not None:
            font = fonts.get(_tag("eastAsia"), fonts.get(_tag("ascii")))
            theme = fonts.get(_tag("eastAsiaTheme"), fonts.get(_tag("asciiTheme")))
            if font:
                add("フォント「" + font + "」")
            elif theme:
                add("フォント(テーマ:" + theme + ")")
        order += 1
        vertical = _val(_find(properties, "vertAlign"))
        if vertical in ("superscript", "subscript"):
            add("上付き" if vertical == "superscript" else "下付き")
        order += 1
        for name, label in (("vanish", "隠し文字"), ("em", "傍点"), ("bdr", "囲み線"), ("spacing", "文字間隔"),
                            ("w", "文字幅"), ("position", "位置"), ("caps", "大文字"), ("smallCaps", "小型英大文字"),
                            ("shadow", "影"), ("outline", "中抜き"), ("emboss", "浮き出し"), ("imprint", "浮き彫り")):
            element = _find(properties, name)
            if element is not None and (name in ("em", "bdr", "spacing", "w", "position") or _on(element)):
                add(label)
            order += 1
        style = self.styles.get(_val(_find(properties, "rStyle")))
        if style:
            add("文字スタイル「" + style.name + "」")
        order += 1
        east = _find(properties, "eastAsianLayout")
        if east is not None:
            if east.get(_tag("combine"), "0") in ("1", "true"):
                add("割注")
            else:
                add("縦中横")
        return labels

    def _section_text(self, section):
        size = _find(section, "pgSz")
        width = size.get(_tag("w"), "11906") if size is not None else "11906"
        height = size.get(_tag("h"), "16838") if size is not None else "16838"
        try:
            dimensions = sorted((int(width), int(height)))
        except ValueError:
            dimensions = [0, 0]
        paper = None
        for name, w, h in (("A3", 16838, 23811), ("A4", 11906, 16838), ("A5", 8391, 11906),
                            ("B4", 14570, 20636), ("B5", 10318, 14570), ("Letter", 12240, 15840), ("Legal", 12240, 20160)):
            if abs(dimensions[0] - w) <= 20 and abs(dimensions[1] - h) <= 20:
                paper = name
                break
        if paper is None:
            paper = _decimal(width, 1440 / 25.4) + "×" + _decimal(height, 1440 / 25.4) + "mm"
        landscape = (size is not None and size.get(_tag("orient")) == "landscape") or dimensions != [0, 0] and _number(width) > _number(height)
        start = _val(_find(section, "type"), "nextPage")
        return paper + (" 横 " if landscape else " 縦 ") + {
            "nextPage": "次のページから", "continuous": "現在の位置から", "oddPage": "奇数ページから",
            "evenPage": "偶数ページから", "nextColumn": "段区切り"}.get(start, start)

    def intrinsic_reasons(self, nodes, part=None, grouped=False):
        """Return intrinsic refusals only; cross-document planner checks are separate."""
        part = part or self.main
        reasons = []
        for root in nodes:
            for node in _walk(root):
                name, ns = _local(node), _namespace(node)
                if not _known_extension(node) and ((ns == W and name not in _KNOWN_WORD) or
                        ns not in (W, M, MC, A, WP, PIC)):
                    # Object namespaces already have specific refusal reasons.
                    object_ns = any(value in ns.lower() for value in ("vml", "office:office", "chart", "diagram", "wordprocessingshape", "wordprocessinggroup", "ink", "model3d"))
                    if not object_ns:
                        reasons.append("未対応の要素 <" + (node.prefix + ":" if node.prefix else "") + name + ">")
                if ns == W and name in ("commentReference", "commentRangeStart", "commentRangeEnd"):
                    reasons.append("コメント")
                elif ns == W and name in ("footnoteReference", "endnoteReference"):
                    reasons.append("脚注・文末脚注")
                if name == "chart" and "chart" in ns:
                    reasons.append("グラフ")
                elif name == "relIds" and "diagram" in ns:
                    reasons.append("SmartArt")
                elif name in ("wsp", "wgp", "txbxContent"):
                    reasons.append("図形・テキストボックス")
                elif ns == "urn:schemas-microsoft-com:vml":
                    reasons.append("VML の図形・画像")
                elif name == "OLEObject" or (ns == W and name == "object"):
                    reasons.append("埋め込みオブジェクト（OLE）")
                elif ns == W and name == "control":
                    reasons.append("ActiveX コントロール")
                elif "ink" in ns.lower() or "model3d" in ns.lower() or name == "contentPart":
                    reasons.append("インク・3D モデル")
                if ns == W and name.startswith(("moveFrom", "moveTo")):
                    reasons.append("変更履歴の移動")
                if ns == W and name == "dataBinding":
                    reasons.append("データにリンクしたコンテンツ コントロール")
                if name in ("repeatingSection", "repeatingSectionItem"):
                    reasons.append("繰り返しセクション")
                if ns == W and name == "sdt" and _find(node, "sdtPr/showingPlcHdr") is not None:
                    placeholder = _val(_find(node, "sdtPr/placeholder/docPart"))
                    if placeholder not in self.glossary_names:
                        reasons.append("プレースホルダー表示中のコンテンツ コントロール")
                if ns == W and name in ("altChunk", "subDoc", "customXml"):
                    reasons.append({"altChunk": "埋め込み文書（altChunk）", "subDoc": "サブ文書", "customXml": "カスタム XML"}[name])
                if node.tag == _tag("sectPr"):
                    parent = node.getparent()
                    top_level = (parent is not None and (parent.tag == _tag("body") or
                                 parent.tag == _tag("pPr") and parent.getparent() is not None
                                 and parent.getparent().getparent() is not None
                                 and parent.getparent().getparent().tag == _tag("body")))
                    if grouped or not top_level:
                        reasons.append("セクション区切りを含むまとまり")
        return _unique(reasons)

    def _fragment(self, node):
        clone = copy.deepcopy(node)
        self._contexts[clone] = self._context(node)
        self._field_prefixes[clone] = {clone: self._field_state(node)}
        return clone

    def _decorate(self, node, before, after):
        clone = self._fragment(node)
        starts, _ = self._context(node)
        def kept(marker):
            name = _local(marker)
            return name != "proofErr" and not (name in ("bookmarkStart", "bookmarkEnd") and
                (marker.get(_tag("name"), "") if name == "bookmarkStart" else starts.get(marker.get(_tag("id"), ""), "")).startswith("_"))
        before, after = [m for m in before if kept(m)], [m for m in after if kept(m)]
        for marker in reversed(before):
            clone.insert(0, copy.deepcopy(marker))
        for marker in after:
            clone.append(copy.deepcopy(marker))
        return clone

    def _paragraph_line(self, node, part, numbered_notes=False):
        return Line("paragraph", self._style_name(node), self.display_text(node, part, numbered_notes),
                    self._format(node, part, plain=True))

    def _section_line(self, node, part):
        return Line("section", "", self._section_text(node), self._format(node, part, detach_sections=False))

    def _block_lines(self, node, part, numbered_notes=False):
        if node.tag == _tag("p"):
            return (self._paragraph_line(node, part, numbered_notes),)
        if node.tag == _tag("tbl"):
            shape = self.table_shape(node, part, numbered_notes)
            return (shape.table_line,) + tuple(line for row in shape.rows for line in (row.line,) + row.cells)
        if node.tag == _tag("sectPr"):
            return (self._section_line(node, part),)
        if node.tag == _tag("sdt"):
            label = _val(_find(node, "sdtPr/tag"), _val(_find(node, "sdtPr/alias")))
            header = self._fragment(node)
            content = _find(header, "sdtContent")
            if content is not None:
                header.remove(content)
            lines = [Line("content-control", label, "", self._format(header, part))]
            inner = _find(node, "sdtContent")
            if inner is not None:
                for child in _selected_children(inner):
                    lines.extend(self._block_lines(child, part, numbered_notes))
            return tuple(lines)
        if _marker(node):
            return ()
        return (Line("other", _local(node), self._block_text(node, part, numbered_notes), self._format(node, part)),)

    def table_shape(self, table, part=None, numbered_notes=False):
        """Render a table and expose its safe slicing signature and child indices."""
        part = part or self.main
        geometry = etree.Element(_tag("table-geometry"))
        self._contexts[geometry] = self._context(table)
        for child in table:
            if child.tag != _tag("tr"):
                geometry.append(_without_members(child, "tr"))
        row_shapes, signature_rows = [], []
        sliceable = True
        for child in table:
            if child.tag not in (_tag("tblPr"), _tag("tblGrid"), _tag("tr")):
                sliceable = False
        row_nodes = list(_members(table, "tr"))
        for row_number, (tr_index, row) in enumerate(row_nodes, 1):
            rowpr = etree.Element(_tag("row-properties"))
            self._contexts[rowpr] = self._context(table)
            for child in row:
                if child.tag != _tag("tc"):
                    rowpr.append(_without_members(child, "tc"))
                    if child.tag not in (_tag("trPr"), _tag("tblPrEx")):
                        sliceable = False
            row_geometry = etree.SubElement(geometry, _tag("row-widths"))
            cells, cell_indices, cell_signature = [], [], []
            for tc_index, cell in _members(row, "tc"):
                cell_number = len(cells) + 1
                pr = _find(cell, "tcPr")
                widths = _find(pr, "tcW")
                holder = etree.SubElement(row_geometry, _tag("cell-width"))
                if widths is not None:
                    holder.append(copy.deepcopy(widths))
                content = self._fragment(cell)
                copy_pr = _find(content, "tcPr")
                if copy_pr is not None:
                    for width in list(copy_pr.findall(_tag("tcW"))):
                        copy_pr.remove(width)
                    if not len(copy_pr) and not copy_pr.attrib:
                        content.remove(copy_pr)
                cells.append(Line("cell", "R" + str(row_number) + "C" + str(cell_number),
                                  self._block_text(cell, part, numbered_notes), self._format(content, part)))
                cell_indices.append(tc_index)
                span, vertical, horizontal = _find(pr, "gridSpan"), _find(pr, "vMerge"), _find(pr, "hMerge")
                cell_signature.append((_integer(_val(span, "1"), 1), None if vertical is None else _val(vertical, "continue"),
                                       None if horizontal is None else _val(horizontal, "continue")))
                if any(_marker(n) or n.tag in (_tag("sdt"), _tag("customXml")) for n in cell if n.tag != _tag("tcPr")):
                    sliceable = False
            pr = _find(row, "trPr")
            signature_rows.append((_integer(_val(_find(pr, "gridBefore"), "0")), _integer(_val(_find(pr, "gridAfter"), "0")), tuple(cell_signature)))
            row_shapes.append(RowShape(Line("row", "R" + str(row_number), "", self._format(rowpr, part)),
                                       tuple(cells), tr_index, tuple(cell_indices)))
        ranges, fields = {}, []
        for element in _walk(table):
            name = _local(element)
            if element.tag in (_tag("sdt"), _tag("customXml"), _tag("cellIns"), _tag("cellDel"), _tag("cellMerge")):
                sliceable = False
            parent = element.getparent()
            if parent is not None and parent.tag == _tag("trPr") and name in ("ins", "del"):
                sliceable = False
            # Find the cell owned directly by this table, ignoring nested cells.
            owner = next((n for n in element.iterancestors() if n.tag == _tag("tc")
                          and n.getparent() is not None and n.getparent().getparent() is table), None)
            if _marker(element) and parent is not None and parent.tag in (_tag("tr"), _tag("tc"), _tag("tbl")):
                sliceable = False
            if element.tag == _tag("bookmarkStart") and element.get(_tag("name"), "").startswith("_"):
                continue
            if name.endswith("Start") or name == "permStart":
                ranges[(name.removesuffix("Start"), element.get(_tag("id"), ""))] = owner
            elif name.endswith("End") or name == "permEnd":
                key = name.removesuffix("End"), element.get(_tag("id"), "")
                if key in ranges and ranges[key] is not owner:
                    sliceable = False
            elif element.tag == _tag("fldChar"):
                mode = element.get(_tag("fldCharType"))
                if mode == "begin":
                    fields.append(owner)
                elif mode == "end" and fields and fields.pop() is not owner:
                    sliceable = False
        if fields:
            sliceable = False
        grid = _find(table, "tblGrid")
        columns = len(grid.findall(_tag("gridCol"))) if grid is not None else max((len(r.cells) for r in row_shapes), default=0)
        signature = (columns, tuple(signature_rows)) if sliceable else None
        line = Line("table", self._style_name(table, "tblStyle"), str(len(row_shapes)) + "×" + str(columns), self._format(geometry, part))
        return TableShape(signature, line, tuple(row_shapes))

    def _units(self, story):
        children = list(story.root)
        blocks = [i for i, node in enumerate(children) if not _marker(node)]
        if not blocks:
            return ()
        owners, attached = {}, {i: [] for i in blocks}
        for index, node in enumerate(children):
            if not _marker(node):
                owners[index] = index
                continue
            before = next((b for b in reversed(blocks) if b < index and children[b].tag != _tag("sectPr")), None)
            after = next((b for b in blocks if b > index), None)
            owner = before if _local(node).endswith("End") else after
            if owner is None or children[owner].tag == _tag("sectPr"):
                owner = before if before is not None else after
            if owner is not None:
                owners[index] = owner
                attached[owner].append(index)
        ranges, fields, intervals = {}, [], []
        starts, _ = self._context(story.root)
        for index, root in enumerate(children):
            owner = owners.get(index, index)
            for node in _walk(root):
                name = _local(node)
                identifier = node.get(_tag("id"), "")
                hidden = (name in ("bookmarkStart", "bookmarkEnd") and
                          (node.get(_tag("name"), "").startswith("_") if name == "bookmarkStart"
                           else starts.get(identifier, "").startswith("_")))
                if _marker(node) and not hidden:
                    if name.endswith("Start"):
                        ranges[(name.removesuffix("Start"), identifier)] = owner
                    elif name.endswith("End"):
                        begin = ranges.pop((name.removesuffix("End"), identifier), owner)
                        if begin != owner:
                            intervals.append((min(begin, owner), max(begin, owner), False))
                if node.tag == _tag("fldChar"):
                    mode = node.get(_tag("fldCharType"))
                    if mode == "begin":
                        fields.append(owner)
                    elif mode == "end" and fields:
                        begin = fields.pop()
                        if not fields and begin != owner:
                            intervals.append((begin, owner, False))
        if fields:
            intervals.append((fields[0], blocks[-1], True))
        merged = []
        for first, last, unclosed in sorted(intervals):
            if merged and first <= merged[-1][1]:
                previous = merged.pop()
                merged.append((previous[0], max(previous[1], last), previous[2] or unclosed))
            else:
                merged.append((first, last, unclosed))
        by_start = {first: (last, unclosed) for first, last, unclosed in merged}
        result, position = [], 0
        while position < len(blocks):
            first = blocks[position]
            last, unclosed = by_start.get(first, (first, False))
            members = []
            while position < len(blocks) and blocks[position] <= last:
                members.append(blocks[position])
                position += 1
            indices = sorted(i for b in members for i in [b] + attached[b])
            nodes = [children[i] for i in indices]
            grouped = len(members) > 1 or unclosed
            lines = []
            render_blocks = []
            shape = None
            for index in members:
                node = children[index]
                before = [children[i] for i in attached[index] if i < index]
                after = [children[i] for i in attached[index] if i > index]
                decorated = self._decorate(node, before, after) if before or after else node
                lines.extend(self._block_lines(decorated, story.part))
                render_blocks.append(decorated)
                if grouped and node.tag == _tag("p"):
                    section = _find(node, "pPr/sectPr")
                    if section is not None:
                        lines.append(self._section_line(section, story.part))
                        render_blocks.append(section)
                if node.tag == _tag("tbl") and not grouped:
                    shape = self.table_shape(decorated, story.part)
            local = _local(children[first])
            kind = "group" if grouped else {"p": "paragraph", "tbl": "table", "sdt": "content-control", "sectPr": "section"}.get(local, "other")
            reasons = list(self.intrinsic_reasons(nodes, story.part, grouped))
            if unclosed:
                reasons.append("閉じていないフィールド")
            ref_nodes = []
            for node in nodes:
                if not grouped and node.tag == _tag("p") and _find(node, "pPr/sectPr") is not None:
                    node = self._fragment(node)
                    pr = _find(node, "pPr")
                    pr.remove(_find(pr, "sectPr"))
                ref_nodes.append(node)
            unit = Unit(kind, min(indices), max(indices) + 1, lines, self.refs(ref_nodes, story.part), reasons,
                        shape, nodes=nodes)
            unit.indices = tuple(indices)
            unit.blocks = tuple(render_blocks)
            result.append(unit)
            if not grouped and children[first].tag == _tag("p"):
                section = _find(children[first], "pPr/sectPr")
                if section is not None:
                    section_unit = Unit("section", first, first + 1, (self._section_line(section, story.part),),
                                        self.refs((section,), story.part), host=first, nodes=(section,))
                    section_unit.indices = (first,)
                    section_unit.blocks = (section,)
                    result.append(section_unit)
        return tuple(result)


def _known_extension(node):
    """Recognize preserved property extensions without allowing object payloads.

    DrawingML extensions are open containers (including picture effects and
    SVG rels). Their descendants still pass the independent object/ref/id
    checks; being in an extension list only suppresses the unknown-element
    reason. Word's Microsoft property extensions are likewise safe to copy.
    """
    name, namespace = _local(node), _namespace(node)
    for ancestor in (node,) + tuple(node.iterancestors()):
        if ancestor.tag == _tag("ext", A):
            parent = ancestor.getparent()
            if parent is not None and parent.tag == _tag("extLst", A):
                return True
    if (namespace == "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
            and name in {"sizeRelH", "sizeRelV", "pctWidth", "pctHeight"}):
        return True
    if re.fullmatch(r"http://schemas[.]microsoft[.]com/office/word/20[0-9]{2}/wordml(?:/[A-Za-z0-9]+)?", namespace):
        properties = {"pPr", "rPr", "tblPr", "trPr", "tcPr", "sdtPr", "sdtEndPr"}
        if any(ancestor.tag in {_tag(value) for value in properties}
               for ancestor in node.iterancestors()):
            return True
    return ((name == "symEx" and namespace.endswith("/symex"))
            or (name == "svgBlip" and namespace.endswith("/SVG/main"))
            or (name == "creationId" and namespace == "http://schemas.microsoft.com/office/drawing/2014/main")
            or ("/word/" in namespace and name in {
                "checkbox", "checked", "checkedState", "uncheckedState", "ligatures", "numForm", "numSpacing", "stylisticSets",
                "stylisticSet", "contextualAlternates", "repeatingSection", "repeatingSectionItem", "sectionTitle", "doNotAllowInsertDeleteSection",
                "collapsed", "scene3d", "props3d", "shadow", "glow", "reflection", "textOutline", "textFill"}))


def _structure(node):
    """Compare normalized XML without namespace declarations or attribute order."""
    if node is None:
        return None
    text = node.text or ""
    return [node.tag, sorted(node.attrib.items()), text if text.strip() else ""] + [_structure(n) for n in node]


def _integer(value, default=0):
    try:
        return int(value)
    except (ValueError, TypeError):
        return default


def _number(value):
    try:
        return float(value)
    except (ValueError, TypeError):
        return 0


def _decimal(value, divisor):
    return format(_number(value) / divisor, ".2f").rstrip("0").rstrip(".") or "0"


def _increment(value):
    try:
        return str(int(value) + 1)
    except (ValueError, TypeError):
        return value


def _excerpt(text):
    return text[:10] + ("…" if len(text) > 10 else "")


def _members(root, name):
    """Yield logical rows/cells through wrappers, retaining the owning child index."""
    wanted = _tag(name)
    def nested(node):
        for child in _selected_children(node):
            if child.tag == wanted:
                yield child
            elif child.tag == _tag("sdt"):
                content = _find(child, "sdtContent")
                if content is not None:
                    yield from nested(content)
            elif child.tag in (_tag("customXml"), _tag("ins"), _tag("moveTo")):
                yield from nested(child)
    for index, child in enumerate(root):
        if child.tag == wanted:
            yield index, child
        elif child.tag == _tag("sdt"):
            content = _find(child, "sdtContent")
            if content is not None:
                yield from ((index, member) for member in nested(content))
        elif child.tag in (_tag("customXml"), _tag("ins"), _tag("moveTo")):
            yield from ((index, member) for member in nested(child))


def _without_members(node, name):
    """Retain wrapper metadata; live row/cell content is owned by other lines."""
    clone = copy.deepcopy(node)
    def strip(element):
        if element.tag in (_tag("del"), _tag("moveFrom")):
            return
        for child in list(element):
            if child.tag == _tag(name):
                element.remove(child)
            else:
                strip(child)
    strip(clone)
    return clone


def _marker(node):
    return _namespace(node) == W and (_local(node).endswith(("Start", "End")) or _local(node) == "proofErr")


def load(data, purpose="edit"):
    """Load a Transitional Word package, applying the documented refusal order."""
    return Document(data, purpose)


def placeholder_lines(data):
    """Return the first render heading without parsing stories or loading images."""
    # The converter has already sniffed the content type. Opening the ZIP here
    # avoids presenting a heading for a truncated request, without full load().
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            if "_rels/.rels" not in archive.namelist():
                return []
        return ["# 本文"]
    except (OSError, ValueError, zipfile.BadZipFile):
        return []


def _safe_format(value, list_markers=False):
    """Escape controls in informational format text, which is outside JSON."""
    text = _json(value)[1:-1].replace('\\"', '"')
    if list_markers:
        # Marker fallbacks are already escaped for display; keep one backslash.
        text = re.sub(r'(?:箇条書き|番号)「.*」L[0-9]+',
                      lambda match: re.sub(r'\\\\u(f0[0-9a-f]{2})', r'\\u\1', match.group()), text)
    return text


def _text_line(line, indent=""):
    suffix = " " + _safe_format(line.format, list_markers=True) if line.format else ""
    text = _json(line.text)
    if line.kind == "paragraph":
        return indent + ("[" + _safe_format(line.label) + "] " if line.label else "") + text + suffix
    if line.kind == "table":
        return indent + "表" + (" [" + _safe_format(line.label) + "]" if line.label else "") + " " + text + suffix
    if line.kind == "row":
        return indent + "  行 " + line.label + " " + text + suffix
    if line.kind == "cell":
        return indent + "  " + line.label + " " + text + suffix
    if line.kind == "section":
        return indent + "セクション " + text + suffix
    if line.kind == "content-control":
        return indent + "コンテンツ コントロール [" + _safe_format(line.label) + "] " + text + suffix
    return indent + _safe_format(line.label) + " " + text + suffix


def render_text(data):
    """Render id-free UTF-8 diff text, including unselectable stories and notes."""
    document = load(data, purpose="view")
    output = []

    def render_node(node, part, indent=""):
        if node.tag == _tag("sdt"):
            lines = document._block_lines(node, part, True)
            if lines:
                output.append(_text_line(lines[0], indent))
            content = _find(node, "sdtContent")
            if content is not None:
                for child in _selected_children(content):
                    render_node(child, part, indent + "  ")
            return
        for line in document._block_lines(node, part, True):
            output.append(_text_line(line, indent))
        if node.tag == _tag("p"):
            for element in _walk(node):
                if element.tag == _tag("txbxContent"):
                    for paragraph in element.findall(_tag("p")):
                        output.append(indent + "  テキストボックス " + _json(document.display_text(paragraph, part, True)))

    for story in document.stories:
        output.append("# " + _safe_format(story.title))
        for unit in story.units:
            for node in unit.blocks:
                render_node(node, story.part)
    for kind, title in (("footnote", "脚注"), ("endnote", "文末脚注"), ("comment", "コメント")):
        notes = sorted(((key, info) for key, info in document.notes.items() if key[0] == kind), key=lambda item: item[1].order)
        if notes:
            output.append("# " + title)
            for key, info in notes:
                prefix = title + " " + str(info.order)
                if kind == "comment":
                    prefix += " " + _safe_format(info.author) + " " + _safe_format(info.date)
                text = document._block_text(document.notes.nodes[key], document.notes.parts[kind], True)
                output.append(prefix + " " + _json(text))
    if document.glossary_part:
        document = copy.copy(document)
        document.main = document.glossary_part
        style_part = document.related_part(document.main, "styles")
        document.styles = Styles(document._tree(style_part).getroot(), style_part) if style_part else Styles()
        document.numbering = Numbering()
        document._load_numbering()
        for node in document._tree(document.glossary_part).getroot().iter(_tag("docPart")):
            name = _val(_find(node, "docPartPr/name"))
            output.append("# 文書パーツ: " + _safe_format(name))
            content = _find(node, "docPartBody")
            if content is not None:
                for child in _selected_children(content):
                    render_node(child, document.glossary_part)
    return "\n".join(output) + "\n"
