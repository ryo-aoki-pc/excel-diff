"""Compare Word slots, apply selected XML pieces and validate the package.

choose(projection="raw_changed") is for staging; projection="changed" is for
merge. last_swaps counts equal pieces implicitly copied for hidden references.
"""

import copy
import hashlib
import io
import json
import difflib
import posixpath
import re
import zipfile
from collections import Counter, namedtuple

from lxml import etree

import word_ooxml as word

W, R, MC, WP = word.W, word.R, word.MC, word.WP
_tag, _local, _namespace = word._tag, word._local, word._namespace
_find, _val = word._find, word._val
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
VALIDATION_ERROR = "作り直した Word 文書の検証に失敗しました。書き込みは行っていません"
_REVISION = {"ins", "del", "cellIns", "cellDel", "cellMerge"}
_STYLE_REFS = {"pStyle", "rStyle", "tblStyle"}
_PREFIX_ATTRS = {"Ignorable", "ProcessContent", "MustUnderstand", "Requires"}


def _serialize(tree, original):
    declaration = original.lstrip().startswith(b"<?xml")
    standalone = re.search(br"<\?xml[^?]*standalone\s*=\s*['\"](yes|no)['\"]", original)
    options = {"encoding": "UTF-8", "xml_declaration": declaration}
    if standalone:
        options["standalone"] = standalone[1] == b"yes"
    result = etree.tostring(tree, **options)
    if original.endswith(b"\n") and not result.endswith(b"\n"):
        result += b"\n"
    return result


def write_zip(base, edited):
    """Keep the original member metadata, order and archive comment."""
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w") as archive:
        for info in base.infos:
            archive.writestr(copy.copy(info), edited.get(info.filename, base.parts[info.filename]))
        for part in sorted(set(edited) - set(base.parts)):
            info = zipfile.ZipInfo(part, (1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            archive.writestr(info, edited[part])
        archive.comment = base.comment
    return stream.getvalue()


def _number(value, default=0):
    try:
        return int(value)
    except (ValueError, TypeError):
        return default


def _elements(nodes):
    for node in nodes:
        yield from node.iter()


def _id_kind(node):
    if node.tag in (_tag("bookmarkStart"), _tag("bookmarkEnd")):
        return "bookmark"
    if node.tag in (_tag("permStart"), _tag("permEnd")):
        return "perm"
    if _namespace(node) == W and (_local(node) in _REVISION or _local(node).endswith("PrChange")):
        return "revision"
    if node.tag == _tag("id") and node.getparent() is not None and node.getparent().tag == _tag("sdtPr"):
        return "sdt"
    if node.tag == _tag("docPr", WP):
        return "docPr"
    return None


def _id_value(node, kind):
    return node.get("id", "") if kind == "docPr" else _val(node) if kind == "sdt" else node.get(_tag("id"), "")


def _package_roots(document):
    """Include orphaned headers and notes, but never the glossary."""
    for part, tree in document._trees.items():
        root = tree.getroot()
        if root.tag in (_tag("document"), _tag("hdr"), _tag("ftr"), _tag("footnotes"), _tag("endnotes"), _tag("comments")):
            yield part, root


class Plan:
    """Mutable, rebuild-local dependency map; reasons() uses an isolated plan."""
    def __init__(self, base, source, scan_ids=True):
        self.base, self.source = base, source
        self.trees, self.new_parts = {}, {}
        self.style_map, self.num_map, self.rel_map = {}, {}, {}
        self.namespace_add, self.ignorable_add = {}, {}
        self.transplants = []
        self.id_maps = {"bookmark": {}, "perm": {}}
        self.used = {kind: set() for kind in ("bookmark", "perm", "revision", "sdt", "docPr")}
        if scan_ids:
            for part, root in _package_roots(base):
                for node in root.iter():
                    kind = _id_kind(node)
                    if kind:
                        self.used[kind].add(_id_value(node, kind))
        self.next_ids = {kind: max((_number(i) for i in used), default=0) + 1
                         for kind, used in self.used.items()}
        self.style_ids = set(base.styles)
        self.num_ids = set(base.numbering)
        self.abstract_ids = set(base.numbering.abstracts)
        self.part_names = {part.casefold() for part in base.parts}

    def _tree(self, part):
        if part not in self.trees:
            self.trees[part] = self.base.fresh_tree(part)
        return self.trees[part]

    def _types(self):
        return self._tree("[Content_Types].xml").getroot()

    def _current_types(self):
        tree = self.trees.get("[Content_Types].xml")
        return word.ContentTypes(tree.getroot()) if tree is not None else self.base.content_types

    def _content_type(self, part, content_type, override=False):
        tree = self.trees.get("[Content_Types].xml")
        root = tree.getroot() if tree is not None else self.base._tree("[Content_Types].xml").getroot()
        extension = posixpath.splitext(part)[1][1:].lower()
        defaults = {n.get("Extension", "").lower(): n.get("ContentType")
                    for n in root if _local(n) == "Default"}
        if override or extension in defaults and defaults[extension] != content_type:
            if not any(_local(n) == "Override" and n.get("PartName", "").lstrip("/").casefold() == part.casefold() for n in root):
                etree.SubElement(self._types(), _tag("Override", CT), PartName="/" + part, ContentType=content_type)
        elif extension not in defaults:
            etree.SubElement(self._types(), _tag("Default", CT), Extension=extension, ContentType=content_type)

    def _style(self, identifier):
        if identifier in self.style_map:
            return self.style_map[identifier]
        style = self.source.styles.get(identifier)
        if style is None:
            raise word.DocumentError("存在しないスタイルの参照")
        key = style.type, style.name.casefold()
        same_name_source = [s for s in self.source.styles.values() if s.name.casefold() == key[1]]
        same_name_base = [s for s in self.base.styles.values() if s.name.casefold() == key[1]]
        if len(same_name_source) > 1 or len(same_name_base) > 1:
            raise word.DocumentError("スタイル名の重複「" + style.name + "」")
        candidates = self.base.styles.by_key.get(key, ())
        if candidates:
            if self.base.styles[candidates[0]].name != style.name:
                raise word.DocumentError("スタイル名の大文字・小文字の違い「" + style.name + "」")
            self.style_map[identifier] = candidates[0]
            return candidates[0]
        if same_name_base:
            raise word.DocumentError("スタイルの種類の違い「" + style.name + "」")
        if not self.base.styles.part:
            raise word.DocumentError("スタイル定義がない文書には取り込めません")
        target = identifier
        counter = 1
        while target in self.style_ids:
            target = identifier + "_" + str(counter)
            counter += 1
        self.style_ids.add(target)
        self.style_map[identifier] = target
        node = copy.deepcopy(self.source.styles.nodes[identifier])
        node.set(_tag("styleId"), target)
        node.attrib.pop(_tag("default"), None)
        node.attrib.pop(_tag("rsid"), None)
        for child in list(node):
            if child.tag == _tag("rsid"):
                node.remove(child)
        # Register before walking dependencies: basedOn/link cycles are legal.
        self._tree(self.base.styles.part).getroot().append(node)
        for element in node.iter():
            if element.tag in (_tag("basedOn"), _tag("next"), _tag("link")) or _local(element) in _STYLE_REFS:
                element.set(_tag("val"), self._style(_val(element)))
            elif element.tag == _tag("numId"):
                element.set(_tag("val"), self._num(_val(element)))
        for element in node.iter():
            for attr, value in list(element.attrib.items()):
                if _namespace(attr) == R:
                    element.set(attr, self._rel(value, self.source.styles.part, self.base.styles.part))
        self._namespaces((node,), self.source.styles.part, self.base.styles.part)
        return target

    def _numbering_root(self):
        part = self.base.numbering.part
        if part:
            return self._tree(part).getroot(), part
        part = "word/numbering.xml"
        if part not in self.trees:
            source_root = self.source._tree(self.source.numbering.part).getroot()
            root = etree.Element(_tag("numbering"), nsmap=source_root.nsmap)
            if source_root.get(_tag("Ignorable", MC)):
                root.set(_tag("Ignorable", MC), source_root.get(_tag("Ignorable", MC)))
            self.trees[part] = etree.ElementTree(root)
            self.part_names.add(part.casefold())
            self._content_type(part, "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml", True)
            self._add_rel(self.base.main, R + "/numbering", posixpath.relpath(part, posixpath.dirname(self.base.main)), False)
        return self.trees[part].getroot(), part

    def _num(self, identifier):
        if identifier == "0":
            return "0"
        if identifier in self.num_map:
            return self.num_map[identifier]
        definition = self.source.numbering.get(identifier)
        if definition is None or definition.abstract_id not in self.source.numbering.abstracts:
            raise word.DocumentError("存在しない番号付けの参照")
        candidates = [i for i, value in self.base.numbering.items() if value.def_hash == definition.def_hash]
        candidates.sort(key=lambda value: (_number(value), value))
        if candidates:
            target = identifier if identifier in candidates else candidates[0]
            self.num_map[identifier] = target
            return target
        if definition.pic_bullet:
            raise word.DocumentError("画像の行頭文字")
        if definition.style_link:
            raise word.DocumentError("リスト スタイル")
        root, part = self._numbering_root()
        target = str(max((_number(i) for i in self.num_ids), default=0) + 1)
        abstract_id = str(max((_number(i) for i in self.abstract_ids), default=-1) + 1)
        self.num_ids.add(target)
        self.abstract_ids.add(abstract_id)
        self.num_map[identifier] = target
        abstract = copy.deepcopy(self.source.numbering.abstracts[definition.abstract_id])
        abstract.set(_tag("abstractNumId"), abstract_id)
        nsid = _find(abstract, "nsid")
        if nsid is None:
            nsid = etree.Element(_tag("nsid"))
            abstract.insert(0, nsid)
        existing_nsids = {_val(n) for n in root.iter(_tag("nsid"))}
        counter = 0
        while True:
            digest = hashlib.sha256((definition.def_hash + abstract_id + (":" + str(counter) if counter else "")).encode()).hexdigest()[:8].upper()
            if digest not in existing_nsids:
                break
            counter += 1
        nsid.set(_tag("val"), digest)
        num = copy.deepcopy(self.source.numbering.nodes[identifier])
        num.set(_tag("numId"), target)
        _find(num, "abstractNumId").set(_tag("val"), abstract_id)
        for container in (abstract, num):
            for element in container.iter():
                if _local(element) in _STYLE_REFS:
                    element.set(_tag("val"), self._style(_val(element)))
        before = next((i for i, n in enumerate(root) if n.tag in (_tag("num"), _tag("numIdMacAtCleanup"))), len(root))
        root.insert(before, abstract)
        before = next((i for i, n in enumerate(root) if n.tag == _tag("numIdMacAtCleanup")), len(root))
        root.insert(before, num)
        self._namespaces((abstract, num), self.source.numbering.part, part)
        return target

    def _rels_root(self, part):
        name = word._rels_path(part)
        if name not in self.trees:
            if name in self.base.parts:
                self.trees[name] = self.base.fresh_tree(name)
            else:
                self.trees[name] = etree.ElementTree(etree.Element(_tag("Relationships", PR), nsmap={None: PR}))
                self._content_type(name, "application/vnd.openxmlformats-package.relationships+xml")
        return self.trees[name].getroot()

    def _add_rel(self, owner, kind, target, external):
        root = self._rels_root(owner)
        identifiers = {node.get("Id", "") for node in root}
        identifier = "rId" + str(max((_number(i[3:]) for i in identifiers if re.fullmatch(r"rId[0-9]+", i)), default=0) + 1)
        while identifier in identifiers:
            identifier = "rId" + str(_number(identifier[3:]) + 1)
        attrs = {"Id": identifier, "Type": kind, "Target": target}
        if external:
            attrs["TargetMode"] = "External"
        etree.SubElement(root, _tag("Relationship", PR), **attrs)
        return identifier

    def _rel_identity(self, rel, document):
        if rel.external:
            return rel.type, True, rel.target
        if rel.path not in document.parts:
            raise word.DocumentError("未対応の参照 " + rel.type.rpartition("/")[2])
        return rel.type, False, document.part_sha(rel.path), document.content_types.resolve(rel.path)

    def _rel(self, identifier, source_part, target_part):
        key = source_part, target_part, identifier
        if key in self.rel_map:
            return self.rel_map[key]
        rel = self.source.rels(source_part).get(identifier)
        if rel is None:
            raise word.DocumentError("未対応の参照 " + identifier)
        suffix = rel.type.rpartition("/")[2]
        if rel.type != R + "/image" and not (rel.type == R + "/hyperlink" and rel.external):
            raise word.DocumentError("未対応の参照 " + suffix)
        identity = self._rel_identity(rel, self.source)
        # Search both original and already-added relationships, in XML order.
        rel_name = word._rels_path(target_part)
        root = self.trees[rel_name].getroot() if rel_name in self.trees else self.base._trees[rel_name].getroot() if rel_name in self.base._trees else ()
        for candidate in root:
            external = candidate.get("TargetMode") == "External"
            if candidate.get("Type") != rel.type or external != rel.external:
                continue
            if external:
                equal = candidate.get("Target") == rel.target
            else:
                destination = word._target(target_part, candidate.get("Target", ""))
                value = self.new_parts.get(destination, self.base.parts.get(destination))
                ctype = self.base.content_types.resolve(destination)
                if destination in self.new_parts:
                    ctype = self._current_types().resolve(destination)
                equal = value is not None and hashlib.sha256(value).hexdigest() == identity[2] and ctype == identity[3]
            if equal:
                self.rel_map[key] = candidate.get("Id")
                return candidate.get("Id")
        if rel.external:
            target = self._add_rel(target_part, rel.type, rel.target, True)
        else:
            if not identity[3]:
                raise word.DocumentError("未対応の参照 " + suffix)
            destination = next((p for p in self.base.parts if p.startswith("word/media/")
                                and self.base.content_types.resolve(p) == identity[3]
                                and self.base.part_sha(p) == identity[2]), None)
            if destination is None:
                destination = next((p for p, value in self.new_parts.items() if p.startswith("word/media/")
                                    and hashlib.sha256(value).hexdigest() == identity[2]
                                    and self._current_types().resolve(p) == identity[3]), None)
            if destination is None:
                extension = posixpath.splitext(rel.path)[1].lower()
                number = 1
                while ("word/media/image" + str(number) + extension).casefold() in self.part_names:
                    number += 1
                destination = "word/media/image" + str(number) + extension
                self.part_names.add(destination.casefold())
                self.new_parts[destination] = self.source.parts[rel.path]
                self._content_type(destination, identity[3])
            target = self._add_rel(target_part, rel.type, posixpath.relpath(destination, posixpath.dirname(target_part)), False)
        self.rel_map[key] = target
        return target

    def _namespaces(self, nodes, source_part, target_part):
        source_root = self.source._tree(source_part).getroot()
        target_root = self.trees[target_part].getroot() if target_part in self.trees else self.base._tree(target_part).getroot()
        used = {}
        def remember(prefix, uri):
            if prefix in used and used[prefix] != uri:
                raise word.DocumentError("名前空間の接頭辞の衝突 " + prefix)
            used[prefix] = uri
        for node in _elements(nodes):
            if not isinstance(node.tag, str):
                continue
            if node.prefix:
                remember(node.prefix, _namespace(node))
            for attr, value in node.attrib.items():
                uri = _namespace(attr)
                if uri and uri != word.XML:
                    # Clark names erase the lexical prefix. XPath retains it,
                    # including aliases for a URI already bound elsewhere.
                    lexical = node.xpath("name(@*[local-name()=$local and namespace-uri()=$uri])", local=_local(attr), uri=uri)
                    prefix = lexical.split(":", 1)[0] if ":" in lexical else None
                    if prefix:
                        remember(prefix, uri)
                if _local(attr) in _PREFIX_ATTRS and (_namespace(attr) == MC or attr == "Requires"):
                    for token in value.split():
                        prefix = token.split(":", 1)[0]
                        if prefix not in node.nsmap:
                            raise word.DocumentError("名前空間の接頭辞の衝突 " + prefix)
                        remember(prefix, node.nsmap[prefix])
        additions = self.namespace_add.setdefault(target_part, {})
        for prefix, uri in used.items():
            current = additions.get(prefix, target_root.nsmap.get(prefix))
            if current is not None and current != uri:
                raise word.DocumentError("名前空間の接頭辞の衝突 " + prefix)
            if current is None:
                additions[prefix] = uri
        ignored = self.ignorable_add.setdefault(target_part, [])
        present = target_root.get(_tag("Ignorable", MC), "").split()
        for prefix in source_root.get(_tag("Ignorable", MC), "").split():
            if source_root.nsmap.get(prefix) in used.values() and prefix not in present and prefix not in ignored:
                ignored.append(prefix)
                uri = source_root.nsmap.get(prefix)
                current = additions.get(prefix, target_root.nsmap.get(prefix))
                if current is not None and current != uri:
                    raise word.DocumentError("名前空間の接頭辞の衝突 " + prefix)
                if current is None:
                    additions[prefix] = uri
        if ignored:
            current = additions.get("mc", target_root.nsmap.get("mc"))
            if current is not None and current != MC:
                raise word.DocumentError("名前空間の接頭辞の衝突 mc")
            if current is None:
                additions["mc"] = MC

    def reasons(self, nodes, source_part, target_part):
        """Return all planner refusals without reserving ids or editing state."""
        trial = Plan(self.base, self.source, scan_ids=False)
        reasons = []
        def check(action):
            try:
                action()
            except word.DocumentError as error:
                reasons.append(str(error))
        for element in _elements(nodes):
            if element.tag in tuple(_tag(n) for n in _STYLE_REFS):
                check(lambda element=element: trial._style(_val(element)))
            elif element.tag == _tag("numId"):
                check(lambda element=element: trial._num(_val(element)))
            elif element.tag == _tag("sdt") and _find(element, "sdtPr/showingPlcHdr") is not None:
                if _val(_find(element, "sdtPr/placeholder/docPart")) not in self.base.glossary_names:
                    reasons.append("プレースホルダー表示中のコンテンツ コントロール")
            for attr, value in element.attrib.items():
                if _namespace(attr) == R:
                    check(lambda value=value: trial._rel(value, source_part, target_part))
        check(lambda: trial._namespaces(nodes, source_part, target_part))
        return tuple(dict.fromkeys(reasons))

    def _strip_noise(self, nodes):
        goback = {n.get(_tag("id"), "") for n in _elements(nodes)
                  if n.tag == _tag("bookmarkStart") and n.get(_tag("name")) == "_GoBack"}
        for node in list(_elements(nodes)):
            for attr in list(node.attrib):
                ns, name = _namespace(attr), _local(attr)
                if ((ns == W and name.startswith("rsid")) or
                        (ns.endswith("/wordml") and name in ("paraId", "textId")) or
                        (ns.endswith("/wordprocessingDrawing") and name in ("anchorId", "editId"))):
                    del node.attrib[attr]
            if (node.tag in (_tag("proofErr"), _tag("lastRenderedPageBreak")) or
                    node.tag in (_tag("bookmarkStart"), _tag("bookmarkEnd")) and node.get(_tag("id")) in goback):
                if node.getparent() is not None:
                    node.getparent().remove(node)

    def _ids(self, nodes, source_part):
        for node in _elements(nodes):
            kind = _id_kind(node)
            if not kind:
                continue
            old = _id_value(node, kind)
            if kind in self.id_maps:
                key = source_part, old
                if key not in self.id_maps[kind]:
                    self.id_maps[kind][key] = str(self.next_ids[kind])
                    self.next_ids[kind] += 1
                value = self.id_maps[kind][key]
            elif kind == "sdt" and old not in self.used[kind]:
                value = old
            elif kind == "sdt":
                number = 1
                while str(number) in self.used[kind]:
                    number += 1
                value = str(number)
            else:
                value = str(self.next_ids[kind])
                self.next_ids[kind] += 1
            self.used[kind].add(value)
            node.set("id" if kind == "docPr" else _tag("val") if kind == "sdt" else _tag("id"), value)

    def transplant(self, nodes, source_part, target_part):
        """Rewrite independent deep copies of source XML, in fix-up order."""
        nodes = tuple(nodes)
        self._strip_noise(nodes)
        for node in _elements(nodes):
            if node.tag in tuple(_tag(n) for n in _STYLE_REFS):
                node.set(_tag("val"), self._style(_val(node)))
        for node in _elements(nodes):
            if node.tag == _tag("numId"):
                node.set(_tag("val"), self._num(_val(node)))
        for node in _elements(nodes):
            for attr, value in list(node.attrib.items()):
                if _namespace(attr) == R:
                    node.set(attr, self._rel(value, source_part, target_part))
        self._ids(nodes, source_part)
        self._namespaces(nodes, source_part, target_part)
        self.transplants.extend((target_part, node) for node in nodes)

    def _incoming(self, node):
        return any(node is root or root in node.iterancestors() for part, root in self.transplants)

    def _repair_hidden(self, root, part):
        starts, ends, names = {}, {}, {}
        for node in root.iter():
            if node.tag == _tag("bookmarkStart"):
                starts.setdefault(node.get(_tag("id"), ""), []).append(node)
                names.setdefault(node.get(_tag("name"), "").casefold(), []).append(node)
            elif node.tag == _tag("bookmarkEnd"):
                ends.setdefault(node.get(_tag("id"), ""), []).append(node)
        remove = []
        for name, group in names.items():
            if len(group) < 2:
                continue
            if not name.startswith("_"):
                raise word.DocumentError("ブックマーク「" + group[0].get(_tag("name"), "") + "」が重複します。元の位置の段落の変更も選んでください。")
            incoming = [node for node in group if self._incoming(node)]
            # Keeping one source copy is deterministic when all copies arrived.
            if len(incoming) == len(group):
                incoming = incoming[1:]
            for node in incoming:
                remove.extend([node] + ends.get(node.get(_tag("id"), ""), []))
        for identifier, group in starts.items():
            if group[0].get(_tag("name"), "").startswith("_") and identifier not in ends:
                remove.extend(group)
        # Identify hidden orphan ends using the origin's names after id mapping.
        hidden_ids = {identifier for identifier, group in starts.items() if group[0].get(_tag("name"), "").startswith("_")}
        hidden_ids.update(identifier for identifier, name in self.base._context(self.base._tree(part).getroot())[0].items() if name.startswith("_"))
        for (source_part, old), identifier in self.id_maps["bookmark"].items():
            source_names = self.source._context(self.source._tree(source_part).getroot())[0]
            if source_names.get(old, "").startswith("_"):
                hidden_ids.add(identifier)
        for identifier, group in ends.items():
            if identifier in hidden_ids and identifier not in starts:
                remove.extend(group)
        for node in dict.fromkeys(remove):
            if node.getparent() is not None:
                node.getparent().remove(node)

    def finish(self, trees):
        """Repair assembled stories, extend root scopes and serialize edits."""
        for part, tree in trees.items():
            self.trees[part] = tree
            self._repair_hidden(tree.getroot(), part)
        for part, additions in self.namespace_add.items():
            if part not in self.trees:
                self.trees[part] = self.base.fresh_tree(part)
            tree = self.trees[part]
            root = tree.getroot()
            ignored = self.ignorable_add.get(part, ())
            if not additions and not ignored:
                continue
            nsmap = dict(root.nsmap)
            nsmap.update(additions)
            new_root = etree.Element(root.tag, nsmap=nsmap)
            for key, value in root.attrib.items():
                new_root.set(key, value)
            if ignored:
                tokens = root.get(_tag("Ignorable", MC), "").split()
                new_root.set(_tag("Ignorable", MC), " ".join(tokens + [p for p in ignored if p not in tokens]))
            new_root.text, new_root.tail = root.text, root.tail
            for child in list(root):
                new_root.append(child)
            tree._setroot(new_root)
        edited = dict(self.new_parts)
        for part, tree in self.trees.items():
            original = self.base.parts.get(part, b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
            value = _serialize(tree, original)
            if part not in self.base.parts or value != self.base.parts[part]:
                edited[part] = value
        return edited


def _assert(condition):
    if not condition:
        raise word.DocumentError(VALIDATION_ERROR)


def _xml_attributes(node):
    return tuple(sorted(node.attrib.items()))


def _structural(node):
    """Namespace declarations and attribute serialization order are immaterial."""
    return (node.tag, tuple(sorted(node.attrib.items())), node.text or "", node.tail or "",
            tuple(_structural(child) for child in node))


def validate(base, source, result, edited, expected, transplants, swapped=()):
    """Validate bytes as written, not the mutable trees used for assembly."""
    try:
        with zipfile.ZipFile(io.BytesIO(result)) as archive:
            names = archive.namelist()
            _assert(len(names) == len(set(names)))
            _assert(set(names) == set(base.parts) | set(edited))
            _assert(names[:len(base.infos)] == [info.filename for info in base.infos])
            _assert(archive.comment == base.comment)
            _assert(not any(p.lower().startswith("_xmlsignatures/") for p in names))
            for part in base.parts:
                if part not in edited:
                    _assert(archive.read(part) == base.parts[part])
        rebuilt = word.load(result, purpose="edit")
        for part in set(rebuilt.parts) - set(base.parts):
            _assert(rebuilt.content_types.resolve(part) is not None)
        _assert([(s.id, s.part, s.roles) for s in rebuilt.stories] == [(s.id, s.part, s.roles) for s in base.stories])
        for story in rebuilt.stories:
            if story.id in expected:
                lines = tuple(line for unit in story.units for line in unit.lines)
                _assert(lines == tuple(expected[story.id]))
        # Equal swaps are allowed only when their display stayed equal.
        for pair in swapped:
            if len(pair) == 2:
                _assert(tuple(pair[0]) == tuple(pair[1]))
        # Every original dependency remains byte-semantically present.
        for part in base._trees:
            if part.endswith(".rels"):
                old = {_xml_attributes(n) for n in base._trees[part].getroot()}
                new = {_xml_attributes(n) for n in rebuilt._tree(part).getroot()}
                _assert(old <= new)
        old_types = {_xml_attributes(n) for n in base._tree("[Content_Types].xml").getroot()}
        new_types = {_xml_attributes(n) for n in rebuilt._tree("[Content_Types].xml").getroot()}
        _assert(old_types <= new_types)
        for part in edited:
            if part.endswith(".rels"):
                owner = "" if part == "_rels/.rels" else posixpath.join(posixpath.dirname(posixpath.dirname(part)), posixpath.basename(part)[:-5])
                for rel in rebuilt.rels(owner).values():
                    if not rel.external:
                        _assert(rel.path in rebuilt.parts and rebuilt.content_types.resolve(rel.path) is not None)
            if part not in rebuilt._trees or part.endswith(".rels") or part == "[Content_Types].xml":
                continue
            root = rebuilt._tree(part).getroot()
            for node in root.iter():
                for attr, value in node.attrib.items():
                    if _namespace(attr) == R:
                        rel = rebuilt.rels(part).get(value)
                        _assert(rel is not None)
                        if not rel.external:
                            _assert(rel.path in rebuilt.parts and rebuilt.content_types.resolve(rel.path) is not None)
                    if _local(attr) in _PREFIX_ATTRS and (_namespace(attr) == MC or attr == "Requires"):
                        _assert(all(token.split(":", 1)[0] in node.nsmap for token in value.split()))
                if node.tag in tuple(_tag(name) for name in _STYLE_REFS):
                    _assert(_val(node) in rebuilt.styles)
                elif node.tag == _tag("numId"):
                    _assert(_val(node) == "0" or _val(node) in rebuilt.numbering)
        for identifier, node in base.styles.nodes.items():
            _assert(identifier in rebuilt.styles.nodes and _structural(node) == _structural(rebuilt.styles.nodes[identifier]))
        for identifier, node in base.numbering.abstracts.items():
            _assert(identifier in rebuilt.numbering.abstracts and _structural(node) == _structural(rebuilt.numbering.abstracts[identifier]))
        for identifier, node in base.numbering.nodes.items():
            _assert(identifier in rebuilt.numbering.nodes and _structural(node) == _structural(rebuilt.numbering.nodes[identifier]))
        defaults = Counter(style.type for style in rebuilt.styles.values() if style.default)
        _assert(all(count <= 1 for count in defaults.values()))
        for identifier, style in rebuilt.styles.items():
            if identifier not in base.styles:
                _assert(all(not value or value in rebuilt.styles for value in (style.based_on, style.next, style.link)))
        if rebuilt.styles.part:
            for node in rebuilt._tree(rebuilt.styles.part).getroot().iter(_tag("numId")):
                _assert(_val(node) == "0" or _val(node) in rebuilt.numbering)
        for definition in rebuilt.numbering.values():
            _assert(definition.abstract_id in rebuilt.numbering.abstracts)
        if rebuilt.numbering.part:
            ranks = {"numPicBullet": 0, "abstractNum": 1, "num": 2, "numIdMacAtCleanup": 3}
            values = [ranks.get(_local(n), -1) for n in rebuilt._tree(rebuilt.numbering.part).getroot()]
            _assert(values == sorted(values) and all(value >= 0 for value in values))
            _assert(values.count(3) <= 1)
        _validate_ids(base, rebuilt, transplants)
        for story in rebuilt.stories:
            if story.id not in expected:
                continue
            _validate_ranges(story.root)
            root = story.root
            if story.id == "body":
                blocks = [n for n in root if n.tag != _tag("sectPr") and not word._marker(n)]
                if not blocks or blocks[-1].tag != _tag("p"):
                    raise word.DocumentError(story.title + " の最後には段落が必要です。最後の段落の変更も選んでください。")
                _assert(_find(blocks[-1], "pPr/sectPr") is None)
            elif not any(not word._marker(n) for n in root):
                raise word.DocumentError(story.title + " が空になります。少なくとも 1 段落を残してください。")
            for cell in root.iter(_tag("tc")):
                blocks = [n for n in cell if n.tag != _tag("tcPr") and not word._marker(n)]
                if not blocks or blocks[-1].tag != _tag("p"):
                    raise word.DocumentError(story.title + " の最後には段落が必要です。最後の段落の変更も選んでください。")
    except word.DocumentError as error:
        if str(error) == VALIDATION_ERROR or "最後には段落が必要です。" in str(error) or "が空になります。" in str(error):
            raise
        raise word.DocumentError(VALIDATION_ERROR) from error
    except (ValueError, KeyError, IndexError, TypeError, OSError, zipfile.BadZipFile, etree.LxmlError) as error:
        raise word.DocumentError(VALIDATION_ERROR) from error


def _validate_ids(base, rebuilt, transplants):
    old = {kind: set() for kind in ("bookmark", "perm", "revision", "sdt", "docPr")}
    counts = {kind: Counter() for kind in old}
    for part, root in _package_roots(base):
        for node in root.iter():
            kind = _id_kind(node)
            if kind:
                old[kind].add(_id_value(node, kind))
    for part, root in _package_roots(rebuilt):
        for node in root.iter():
            kind = _id_kind(node)
            if kind and node.tag not in (_tag("bookmarkEnd"), _tag("permEnd")):
                counts[kind][_id_value(node, kind)] += 1
    # Inspect the actual serialized locations through the rewritten ids. A
    # dropped transplant is independently detected by the content check.
    seen = {kind: set() for kind in old}
    for part, root in transplants:
        for node in root.iter():
            kind = _id_kind(node)
            if not kind or node.tag in (_tag("bookmarkEnd"), _tag("permEnd")):
                continue
            identifier = _id_value(node, kind)
            # A removed element may retain its old getroottree() identity in
            # lxml. Hidden-bookmark repair is allowed to drop an entire root.
            if node.tag == _tag("bookmarkStart") and node.get(_tag("name"), "").startswith("_") and counts[kind][identifier] == 0:
                continue
            _assert(identifier not in old[kind])
            _assert(identifier not in seen[kind] and counts[kind][identifier] == 1)
            seen[kind].add(identifier)


def _validate_ranges(root):
    for start_name, end_name in (("bookmarkStart", "bookmarkEnd"), ("permStart", "permEnd")):
        opened, finished = set(), set()
        for node in word._walk(root):
            identifier = node.get(_tag("id"), "")
            if node.tag == _tag(start_name):
                _assert(identifier not in opened and identifier not in finished)
                opened.add(identifier)
            elif node.tag == _tag(end_name):
                _assert(identifier in opened)
                opened.remove(identifier)
                finished.add(identifier)
        _assert(not opened)
    fields = []
    for node in word._walk(root):
        if node.tag != _tag("fldChar"):
            continue
        mode = node.get(_tag("fldCharType"))
        if mode == "begin":
            fields.append(False)
        elif mode == "separate":
            _assert(fields and not fields[-1])
            fields[-1] = True
        elif mode == "end":
            _assert(fields)
            fields.pop()
    _assert(not fields)

# Alignment and selection deliberately use only ordered sequences.

Record = namedtuple("Record", "id kind label text format")
Unsupported = namedtuple("Unsupported", "story_id story_title slot excerpt reason")
StoryView = namedtuple("StoryView", "id title baseline changed raw_changed")
_RECORD_ID = re.compile(r"[1-9][0-9]*(?:/[1-9][0-9]*| R[1-9][0-9]*(?:C[1-9][0-9]*)?)?\Z")
_KINDS = {"paragraph", "table", "row", "cell", "section", "content-control", "other", "blank"}


def parse_record(line):
    """Validate the complete picker record, including position and padding."""
    try:
        fields = json.loads(line)
        if (not isinstance(fields, list) or len(fields) != 5
                or not all(isinstance(value, str) for value in fields)
                or not _RECORD_ID.fullmatch(fields[0]) or fields[1] not in _KINDS
                or fields[1] == "blank" and any(fields[2:])):
            raise ValueError()
        return Record(*fields)
    except (ValueError, TypeError):
        raise word.DocumentError("Word の選択の行が不正です") from None


def _similar(left, right):
    if left.kind != right.kind:
        return 0
    a, b = left.sim_text, right.sim_text
    if a == b:
        return 1.0
    matcher = difflib.SequenceMatcher(None, a, b, autojunk=False)
    if matcher.real_quick_ratio() < .5 or matcher.quick_ratio() < .5:
        return 0
    value = matcher.ratio() if len(a) + len(b) <= 2000 else matcher.quick_ratio()
    return value if value >= .5 else 0


def _pair_region(left, right, budget):
    n, m = len(left), len(right)
    if n * m > 2500 or n * m > budget[0]:
        if n == m and all(a.kind == b.kind for a, b in zip(left, right)):
            return list(zip(left, right))
        return [(a, None) for a in left] + [(None, b) for b in right]
    budget[0] -= n * m
    best = [[0.0] * (m + 1) for _ in range(n + 1)]
    route = [[0] * (m + 1) for _ in range(n + 1)]
    for i in range(1, n + 1):
        for j in range(1, m + 1):
            similarity = _similar(left[i - 1], right[j - 1])
            diagonal = best[i - 1][j - 1] + similarity if similarity >= .5 else -1
            up, across = best[i - 1][j], best[i][j - 1]
            if diagonal >= up and diagonal >= across:
                best[i][j], route[i][j] = diagonal, 1
            elif up >= across:
                best[i][j], route[i][j] = up, 2
            else:
                best[i][j], route[i][j] = across, 3
    matches, i, j = [], n, m
    while i and j:
        direction = route[i][j]
        if direction == 1:
            matches.append((i - 1, j - 1))
            i, j = i - 1, j - 1
        elif direction == 2:
            i -= 1
        else:
            j -= 1
    result, i, j = [], 0, 0
    for a, b in reversed(matches):
        result.extend((item, None) for item in left[i:a])
        result.extend((None, item) for item in right[j:b])
        result.append((left[a], right[b]))
        i, j = a + 1, b + 1
    result.extend((item, None) for item in left[i:])
    result.extend((None, item) for item in right[j:])
    return result


def _align(left, right, budget):
    prefix = 0
    while prefix < min(len(left), len(right)) and left[prefix].key == right[prefix].key:
        prefix += 1
    suffix = 0
    while (suffix < min(len(left), len(right)) - prefix
           and left[len(left) - suffix - 1].key == right[len(right) - suffix - 1].key):
        suffix += 1
    a = left[prefix:len(left) - suffix if suffix else len(left)]
    b = right[prefix:len(right) - suffix if suffix else len(right)]
    result = list(zip(left[:prefix], right[:prefix]))
    matcher = difflib.SequenceMatcher(None, [u.key for u in a], [u.key for u in b],
                                      autojunk=len(a) * len(b) > 4000000)
    for tag, i, end_i, j, end_j in matcher.get_opcodes():
        if tag == "equal":
            result.extend(zip(a[i:end_i], b[j:end_j]))
        elif tag == "delete":
            result.extend((item, None) for item in a[i:end_i])
        elif tag == "insert":
            result.extend((None, item) for item in b[j:end_j])
        else:
            result.extend(_pair_region(a[i:end_i], b[j:end_j], budget))
    if suffix:
        result.extend(zip(left[-suffix:], right[-suffix:]))
    return result


def _align_lines(left, right):
    result = []
    matcher = difflib.SequenceMatcher(None, [v.key for v in left], [v.key for v in right], autojunk=False)
    for tag, i, end_i, j, end_j in matcher.get_opcodes():
        if tag == "equal":
            result.extend(zip(left[i:end_i], right[j:end_j]))
        else:
            a, b = left[i:end_i], right[j:end_j]
            result.extend((a[k] if k < len(a) else None, b[k] if k < len(b) else None)
                          for k in range(max(len(a), len(b))))
    return result


def _root(tree, story_id):
    return tree.getroot().find(word._tag("body")) if story_id == "body" else tree.getroot()


def _unit_nodes(unit, root=None):
    if unit is None:
        return []
    if unit.kind == "section":
        if root is None:
            return list(unit.nodes)
        return [word._find(root[unit.host], "pPr/sectPr")] if unit.host is not None else [root[unit.start]]
    if root is None:
        nodes = [copy.deepcopy(node) for node in unit.nodes]
    else:
        nodes = [root[i] for i in unit.indices]
    # A paragraph's section belongs to the following section slot.
    for node in nodes:
        if node.tag == word._tag("p"):
            pr = word._find(node, "pPr")
            section = word._find(pr, "sectPr")
            if section is not None:
                pr.remove(section)
    if root is None:
        goback = {n.get(word._tag("id")) for n in nodes if n.tag == word._tag("bookmarkStart")
                  and n.get(word._tag("name")) == "_GoBack"}
        nodes = [n for n in nodes if n.tag != word._tag("proofErr")
                 and not (n.tag in (word._tag("bookmarkStart"), word._tag("bookmarkEnd"))
                          and n.get(word._tag("id")) in goback)]
    return nodes


def _table_node(unit, root=None):
    nodes = unit.nodes if root is None else [root[i] for i in unit.indices]
    return next(node for node in nodes if node.tag == word._tag("tbl"))


def _piece_nodes(unit, piece):
    if unit is None:
        return []
    if piece is None:
        nodes = _unit_nodes(unit)
        if unit.kind == "section":
            nodes = [copy.deepcopy(node) for node in nodes]
            for section in nodes:
                for child in list(section):
                    if word._local(child) in ("headerReference", "footerReference", "printerSettings"):
                        section.remove(child)
        return nodes
    table = _table_node(unit)
    kind, ri, ci = piece
    if kind == "table":
        nodes = [copy.deepcopy(n) for n in table if n.tag != word._tag("tr")]
        nodes.extend(copy.deepcopy(n) for n in table.iter(word._tag("tcW")))
        return nodes
    row = table[unit.table.rows[ri].tr_index]
    if kind == "row":
        return [copy.deepcopy(n) for n in row if n.tag != word._tag("tc")]
    cell = copy.deepcopy(row[unit.table.rows[ri].tc_indices[ci]])
    pr = word._find(cell, "tcPr")
    if pr is not None:
        for width in list(pr.findall(word._tag("tcW"))):
            pr.remove(width)
    return [cell]


class _Slot:
    def __init__(self, story, number, base, source, piece=None, lines=None):
        self.story, self.base, self.source, self.piece = story, base, source, piece
        self.key = number
        if lines is None:
            a, b = base.lines if base else (), source.lines if source else ()
            pairs = _align_lines(a, b)
            multi = max(len(a), len(b)) > 1 or (base and base.kind in ("group", "content-control")) or (source and source.kind in ("group", "content-control"))
        else:
            pairs, multi = [lines], False
        self.base_lines, self.source_lines = [], []
        for i, (a, b) in enumerate(pairs, 1):
            identifier = number + ("/" + str(i) if multi else "")
            self.base_lines.append(word.record_json((identifier,) + tuple(a or word.Line("blank", "", "", ""))))
            self.source_lines.append(word.record_json((identifier,) + tuple(b or word.Line("blank", "", "", ""))))
        self.changed = self.base_lines != self.source_lines
        self.reasons = ()


class Comparison:
    """Aligned, position-checked projections and a validated partial rebuild."""
    def __init__(self, base, source):
        self.base, self.source = base, source
        self.last_swaps = 0
        self._slots, self._by_story, self._paired = {}, {}, []
        views, unsupported = [], []
        source_stories = {story.id: story for story in source.stories}
        budget = [20000]
        self.differences = 0
        self.selectable = 0
        planner = Plan(base, source)
        for a in base.stories:
            b = source_stories.get(a.id)
            if b is None:
                unsupported.append(Unsupported(a.id, a.title, "", "", "ヘッダー・フッターの追加・削除"))
                continue
            self._paired.append((a, b))
            slots = []
            for index, (left, right) in enumerate(_align(a.units, b.units, budget), 1):
                number = str(index)
                if (left and right and left.kind == right.kind == "table"
                        and left.table.signature is not None and left.table.signature == right.table.signature):
                    slots.append(_Slot(a, number, left, right, ("table", None, None),
                                       (left.table.table_line, right.table.table_line)))
                    for ri, (ar, br) in enumerate(zip(left.table.rows, right.table.rows)):
                        row_key = number + " R" + str(ri + 1)
                        slots.append(_Slot(a, row_key, left, right, ("row", ri, None), (ar.line, br.line)))
                        for ci, (ac, bc) in enumerate(zip(ar.cells, br.cells)):
                            slots.append(_Slot(a, row_key + "C" + str(ci + 1), left, right, ("cell", ri, ci), (ac, bc)))
                else:
                    slots.append(_Slot(a, number, left, right))
            baseline, changed, raw = [], [], []
            for slot in slots:
                self._slots[(a.id, slot.key)] = slot
                if slot.changed:
                    slot.reasons = self._reasons(slot, a, b, planner)
                    self.differences += 1
                    if slot.reasons:
                        text = next((parse_record(line).text for line in slot.base_lines + slot.source_lines if parse_record(line).text), "")
                        unsupported.append(Unsupported(a.id, a.title, slot.key, text[:20], "・".join(slot.reasons)))
                    else:
                        self.selectable += 1
                baseline.extend(slot.base_lines)
                raw.extend(slot.source_lines)
                changed.extend(slot.base_lines if slot.reasons else slot.source_lines)
            self._by_story[a.id] = slots
            views.append(StoryView(a.id, a.title, "\n".join(baseline) + ("\n" if baseline else ""),
                                   "\n".join(changed) + ("\n" if changed else ""),
                                   "\n".join(raw) + ("\n" if raw else "")))
        for story in source.stories:
            if story.id not in self._by_story and story.id not in {s.id for s in base.stories}:
                unsupported.append(Unsupported(story.id, story.title, "", "", "ヘッダー・フッターの追加・削除"))
        self.stories, self.unsupported = tuple(views), tuple(unsupported)

    def _reasons(self, slot, a, b, planner, both_planners=False):
        reasons = []
        for doc, unit, story in ((self.base, slot.base, a), (self.source, slot.source, b)):
            if unit is None:
                continue
            nodes = _piece_nodes(unit, slot.piece)
            if unit.kind != "section":
                reasons.extend(reason for reason in doc.intrinsic_reasons(nodes, story.part, unit.kind == "group")
                               if reason != "プレースホルダー表示中のコンテンツ コントロール")
            if slot.piece is None:
                reasons.extend(reason for reason in unit.reasons
                               if reason != "プレースホルダー表示中のコンテンツ コントロール")
            # Placeholder availability is determined by the destination glossary.
            for node in nodes:
                for control in node.iter(word._tag("sdt")):
                    if word._find(control, "sdtPr/showingPlcHdr") is not None:
                        placeholder = word._val(word._find(control, "sdtPr/placeholder/docPart"))
                        if placeholder not in self.base.glossary_names:
                            reasons.append("プレースホルダー表示中のコンテンツ コントロール")
        if (slot.base and slot.base.kind == "section") or (slot.source and slot.source.kind == "section"):
            if not slot.base or not slot.source:
                reasons.append("セクション区切りの追加・削除")
            else:
                def roles(unit):
                    return [(word._local(n), n.get(word._tag("type"), "default")) for n in unit.nodes[0]
                            if word._local(n) in ("headerReference", "footerReference")]
                if roles(slot.base) != roles(slot.source):
                    reasons.append("ヘッダー・フッターの構成")
        if slot.source:
            reasons.extend(planner.reasons(_piece_nodes(slot.source, slot.piece), b.part, a.part))
        if both_planners and slot.base:
            reasons.extend(Plan(self.source, self.base).reasons(_piece_nodes(slot.base, slot.piece), a.part, b.part))
        return tuple(dict.fromkeys(reasons))

    def choose(self, story_id, staged_lines, projection="raw_changed"):
        """Choose only complete records from the projection actually displayed."""
        view = next((v for v in self.stories if v.id == story_id), None)
        if view is None or projection not in ("raw_changed", "changed"):
            raise word.DocumentError("Word の選択の行が不正です")
        baseline = view.baseline.splitlines()
        shown = getattr(view, projection).splitlines()
        staged = [line.removesuffix("\r") for line in staged_lines]
        if len(staged) != len(baseline):
            last = parse_record(baseline[-1]).id if baseline else "1"
            raise word.DocumentError(view.title + " の行数が変わっています（" + last + " 付近）。行全体を選んでください。")
        for a, b, line in zip(baseline, shown, staged):
            if line != a and line != b:
                raise word.DocumentError(view.title + " " + parse_record(a).id + " の選択が不正です。取り込み先か取り込み元の行全体を選んでください。")
        selected, offset = [], 0
        for slot in self._by_story[story_id]:
            length = len(slot.base_lines)
            indices = [i for i in range(offset, offset + length) if baseline[i] != shown[i]]
            takes = [staged[i] == shown[i] for i in indices]
            if takes and all(takes):
                selected.append((story_id, slot.key))
            elif any(takes):
                raise word.DocumentError(view.title + " " + slot.key + " の選択が不完全です。このまとまりの変更行をすべて選んでください。")
            offset += length
        return frozenset(selected)

    def rebuild(self, selected):
        self.last_swaps = 0
        try:
            selected = frozenset(selected)
        except (TypeError, ValueError):
            raise word.DocumentError("Word の選択の行が不正です") from None
        planner = Plan(self.base, self.source)
        if any(not isinstance(key, tuple) or len(key) != 2 or not all(isinstance(v, str) for v in key)
               or key not in self._slots or not self._slots[key].changed for key in selected):
            raise word.DocumentError("Word の選択の行が不正です")
        for key, slot in self._slots.items():
            if key not in selected:
                continue
            a, b = next(pair for pair in self._paired if pair[0].id == key[0])
            reasons = self._reasons(slot, a, b, planner)
            if reasons:
                raise word.DocumentError(a.title + " " + slot.key + " は段落単位で反映できません: " + "・".join(reasons))
        if not selected:
            return self.base.data
        effective, origin_refs = self._equal_swaps(selected, planner)
        trees, expected = self._assemble(effective, planner)
        self._repair_all_duplicates(trees, expected, planner, effective, origin_refs)
        edited = planner.finish(trees)
        result = write_zip(self.base, edited)
        validate(self.base, self.source, result, edited, expected, planner.transplants,
                 tuple((slot.base_lines, slot.source_lines) for key, slot in self._slots.items() if key in effective and key not in selected))
        actual = word.load(result, purpose="edit")
        self._check_refs(origin_refs, actual.all_bookmarks)
        self.last_swaps = len(effective - selected)
        return result

    @staticmethod
    def _reference_error(kind, name):
        raise word.DocumentError(kind + "の参照先「" + name + "」が見つからなくなります。参照先の段落の変更も選んでください。")

    def _check_refs(self, refs, bookmarks):
        for doc, kind, name in refs:
            if name.casefold() in doc.all_bookmarks and name.casefold() not in bookmarks:
                self._reference_error(kind, name)

    def _choices(self, selected):
        for a, b in self._paired:
            for slot in self._by_story[a.id]:
                incoming = (a.id, slot.key) in selected
                doc, story, unit = (self.source, b, slot.source) if incoming else (self.base, a, slot.base)
                nodes = _piece_nodes(unit, slot.piece)
                yield slot, doc, story, doc.refs(nodes, story.part)

    def _equal_swaps(self, selected, planner):
        effective = set(selected)
        # Notes and stories that have no paired projection always remain base.
        paired_ids = set(self._by_story)
        fixed_names, fixed_refs = [], []
        for identifier, title, part, root in self.base.reference_roots:
            if identifier not in paired_ids:
                refs = self.base.refs((root,), part)
                fixed_names.extend(name.casefold() for name in refs.bookmarks)
                fixed_refs.extend((self.base, kind, name) for kind, name in refs.ref_names)
        while True:
            choices = list(self._choices(effective))
            names = set(fixed_names)
            refs = list(fixed_refs)
            for slot, doc, story, info in choices:
                names.update(name.casefold() for name in info.bookmarks)
                refs.extend((doc, kind, name) for kind, name in info.ref_names)
            missing = [(doc, kind, name) for doc, kind, name in refs
                       if name.casefold() in doc.all_bookmarks and name.casefold() not in names]
            if not missing:
                return frozenset(effective), refs
            progress = False
            for doc, kind, name in missing:
                if doc is not self.source or not name.startswith("_"):
                    self._reference_error(kind, name)
                candidates = []
                target_id = None
                target_story = None
                for a, b in self._paired:
                    for slot in self._by_story[a.id]:
                        if (a.id, slot.key) in effective or not slot.base or not slot.source or slot.changed:
                            continue
                        nodes = _piece_nodes(slot.source, slot.piece)
                        start = next((n for root in nodes for n in root.iter(word._tag("bookmarkStart"))
                                      if n.get(word._tag("name"), "").casefold() == name.casefold()), None)
                        if start is not None:
                            candidates.append((a.id, slot.key))
                            target_id, target_story = start.get(word._tag("id")), a.id
                            break
                    if candidates:
                        break
                if not candidates:
                    self._reference_error(kind, name)
                # A loose hidden range may span several equal pieces. Both
                # endpoints must arrive under the source story's single id map.
                found_end = False
                for a, b in self._paired:
                    if a.id != target_story:
                        continue
                    for slot in self._by_story[a.id]:
                        nodes = _piece_nodes(slot.source, slot.piece)
                        if any(n.get(word._tag("id")) == target_id for root in nodes
                               for n in root.iter(word._tag("bookmarkEnd"))):
                            found_end = True
                            key = a.id, slot.key
                            if key not in effective and key not in candidates:
                                if slot.changed or not slot.base:
                                    self._reference_error(kind, name)
                                candidates.append(key)
                    for key in candidates:
                        slot = self._slots[key]
                        if self._reasons(slot, a, b, planner, both_planners=True):
                            self._reference_error(kind, name)
                if not found_end:
                    self._reference_error(kind, name)
                tentative = effective | set(candidates)
                kept_refs, future_names = list(fixed_refs), set(fixed_names)
                remaining_refs = list(fixed_refs)
                for slot, origin, story, info in self._choices(tentative):
                    future_names.update(n.casefold() for n in info.bookmarks)
                    remaining_refs.extend((origin, k, n) for k, n in info.ref_names)
                    if origin is self.base:
                        kept_refs.extend((origin, k, n) for k, n in info.ref_names)
                for key in candidates:
                    slot = self._slots[key]
                    base_names = self.base.refs(_piece_nodes(slot.base, slot.piece), slot.story.part).bookmarks
                    source_story = next(b for a, b in self._paired if a.id == key[0])
                    same_piece = {n.casefold() for n in self.source.refs(_piece_nodes(slot.source, slot.piece), source_story.part).bookmarks}
                    for old in base_names:
                        if old.startswith("_") and old.casefold() not in same_piece:
                            for origin, ref_kind, ref_name in remaining_refs:
                                if ref_name.casefold() == old.casefold():
                                    self._reference_error(ref_kind, ref_name)
                self._check_refs(kept_refs, future_names)
                effective.update(candidates)
                progress = True
                # Recompute after each swap so subsequent guards see its refs.
                break
            if not progress:
                self._check_refs(refs, names)

    def _assemble(self, selected, planner):
        trees, expected = {}, {}
        for a, b in self._paired:
            slots = self._by_story[a.id]
            if not any((a.id, slot.key) in selected for slot in slots):
                continue
            tree = self.base.fresh_tree(a.part)
            root = _root(tree, a.id)
            original = list(root)
            # Resolve all section nodes before detaching any paragraph section.
            base_nodes = {}
            for unit in a.units:
                if unit.kind == "section":
                    base_nodes[unit] = [word._find(original[unit.host], "pPr/sectPr")] if unit.host is not None else [original[unit.start]]
                else:
                    base_nodes[unit] = [original[i] for i in unit.indices]
            items, done_tables = [], set()
            expected_lines = []
            for slot in slots:
                incoming = (a.id, slot.key) in selected
                for line in slot.source_lines if incoming else slot.base_lines:
                    record = parse_record(line)
                    if record.kind != "blank":
                        expected_lines.append(word.Line(*record[1:]))
                if slot.piece is not None:
                    if slot.base in done_tables:
                        continue
                    done_tables.add(slot.base)
                    nodes = base_nodes[slot.base]
                    table = next(n for n in nodes if n.tag == word._tag("tbl"))
                    related = [s for s in slots if s.base is slot.base and s.piece is not None]
                    self._replace_table(table, related, selected, a, b, planner)
                    items.append(("blocks", nodes, slot.key))
                    continue
                unit = slot.source if incoming else slot.base
                if unit is None:
                    continue
                if unit.kind == "section":
                    section = copy.deepcopy(slot.source.nodes[0]) if incoming else base_nodes[unit][0]
                    if incoming:
                        preserved = [copy.deepcopy(n) for n in base_nodes[slot.base][0]
                                     if word._local(n) in ("headerReference", "footerReference", "printerSettings")]
                        for child in list(section):
                            if word._local(child) in ("headerReference", "footerReference", "printerSettings"):
                                section.remove(child)
                        planner.transplant([section], b.part, a.part)
                        for child in reversed([n for n in preserved if word._local(n) != "printerSettings"]):
                            section.insert(0, child)
                        for child in preserved:
                            if word._local(child) == "printerSettings":
                                section.append(child)
                    items.append(("section", [section], slot.key))
                else:
                    nodes = _unit_nodes(unit) if incoming else base_nodes[unit]
                    if incoming:
                        planner.transplant(nodes, b.part, a.part)
                    items.append(("blocks", nodes, slot.key))
            # Detach only after the section items retain their exact elements.
            for kind, nodes, key in items:
                if kind != "blocks":
                    continue
                for node in nodes:
                    if node.tag == word._tag("p"):
                        pr = word._find(node, "pPr")
                        section = word._find(pr, "sectPr")
                        if section is not None:
                            pr.remove(section)
            final_section = max((i for i, item in enumerate(items) if item[0] == "section"), default=-1)
            output = []
            for index, (kind, nodes, key) in enumerate(items):
                if kind != "section":
                    output.extend(nodes)
                elif index == final_section:
                    preceding = next((n for n in reversed(output) if not word._marker(n)), None)
                    if preceding is not None and preceding.tag == word._tag("p") and word._find(preceding, "pPr/sectPr") is not None:
                        raise word.DocumentError("セクション区切りの前に段落が必要です（" + a.title + " " + key + "）。直前の段落の変更も選んでください。")
                    output.extend(nodes)
                else:
                    preceding = next((n for n in reversed(output) if not word._marker(n)), None)
                    if (preceding is None or preceding.tag != word._tag("p")
                            or word._find(preceding, "pPr/sectPr") is not None):
                        raise word.DocumentError("セクション区切りの前に段落が必要です（" + a.title + " " + key + "）。直前の段落の変更も選んでください。")
                    pr = word._find(preceding, "pPr")
                    if pr is None:
                        pr = etree.Element(word._tag("pPr"))
                        preceding.insert(0, pr)
                    following = word._find(pr, "pPrChange")
                    pr.insert(pr.index(following) if following is not None else len(pr), nodes[0])
            root[:] = output
            trees[a.part] = tree
            expected[a.id] = tuple(expected_lines)
        return trees, expected

    def _replace_table(self, table, slots, selected, a, b, planner):
        source_table = _table_node(slots[0].source)
        for slot in slots:
            if (a.id, slot.key) not in selected:
                continue
            kind, ri, ci = slot.piece
            if kind == "table":
                replacements = [copy.deepcopy(n) for n in source_table if n.tag != word._tag("tr")]
                planner.transplant(replacements, b.part, a.part)
                for node in list(table):
                    if node.tag != word._tag("tr"):
                        table.remove(node)
                for index, node in enumerate(replacements):
                    table.insert(index, node)
                for target_row, source_row in zip(table.findall(word._tag("tr")), source_table.findall(word._tag("tr"))):
                    for target_cell, source_cell in zip(target_row.findall(word._tag("tc")), source_row.findall(word._tag("tc"))):
                        source_width = word._find(source_cell, "tcPr/tcW")
                        target_pr = word._find(target_cell, "tcPr")
                        if target_pr is not None:
                            for width in list(target_pr.findall(word._tag("tcW"))):
                                target_pr.remove(width)
                        if source_width is not None:
                            if target_pr is None:
                                target_pr = etree.Element(word._tag("tcPr"))
                                target_cell.insert(0, target_pr)
                            target_pr.insert(0, copy.deepcopy(source_width))
            elif kind == "row":
                ar, br = slot.base.table.rows[ri], slot.source.table.rows[ri]
                target_row, source_row = table.findall(word._tag("tr"))[ri], source_table.findall(word._tag("tr"))[ri]
                nodes = [copy.deepcopy(n) for n in source_row if n.tag != word._tag("tc")]
                planner.transplant(nodes, b.part, a.part)
                for node in list(target_row):
                    if node.tag != word._tag("tc"):
                        target_row.remove(node)
                for index, node in enumerate(nodes):
                    target_row.insert(index, node)
            else:
                ar, br = slot.base.table.rows[ri], slot.source.table.rows[ri]
                target_row, source_row = table.findall(word._tag("tr"))[ri], source_table.findall(word._tag("tr"))[ri]
                old, source_cell = target_row.findall(word._tag("tc"))[ci], source_row.findall(word._tag("tc"))[ci]
                cell = copy.deepcopy(source_cell)
                pr = word._find(cell, "tcPr")
                if pr is not None:
                    for width in list(pr.findall(word._tag("tcW"))):
                        pr.remove(width)
                width = word._find(old, "tcPr/tcW")
                planner.transplant([cell], b.part, a.part)
                if width is not None:
                    if pr is None:
                        pr = etree.Element(word._tag("tcPr"))
                        cell.insert(0, pr)
                    pr.insert(0, copy.deepcopy(width))
                target_row.replace(old, cell)

    def _repair_all_duplicates(self, trees, expected, planner, selected, origin_refs):
        # Bookmark names are package-wide even when their ranges are local.
        roots = {}
        for identifier, title, part, root in self.base.reference_roots:
            roots[part] = (identifier, _root(trees[part], identifier) if part in trees else root)
        names = {}
        for part, (identifier, root) in roots.items():
            for node in root.iter(word._tag("bookmarkStart")):
                names.setdefault(node.get(word._tag("name"), "").casefold(), []).append((part, node))
        transplanted = {n for part, node in planner.transplants for n in node.iter()}
        needed = {name.casefold() for origin, kind, name in origin_refs if origin is self.source}
        base_needed = {name.casefold() for origin, kind, name in origin_refs if origin is self.base}
        for name, entries in names.items():
            if len(entries) < 2:
                continue
            label = entries[0][1].get(word._tag("name"), "")
            def refuse():
                raise word.DocumentError("ブックマーク「" + label + "」が重複します。元の位置の段落の変更も選んでください。")
            if not name.startswith("_"):
                refuse()
            incoming = [(part, node) for part, node in entries if node in transplanted]
            kept = [(part, node) for part, node in entries if node not in transplanted]
            if name in needed and incoming and kept:
                if name in base_needed:
                    refuse()
                dropping = kept
            else:
                dropping = incoming if kept else incoming[1:]
            for part, node in dropping:
                identifier, root = roots[part]
                if part not in trees:
                    story = next((s for s in self.base.stories if s.id == identifier), None)
                    if story is None:
                        refuse()
                    tree = self.base.fresh_tree(part)
                    root = _root(tree, identifier)
                    trees[part] = tree
                    roots[part] = identifier, root
                    expected[identifier] = tuple(line for unit in story.units for line in unit.lines)
                    node = next(n for n in root.iter(word._tag("bookmarkStart"))
                                if n.get(word._tag("name"), "").casefold() == name)
                ident = node.get(word._tag("id"))
                for element in list(root.iter()):
                    if element is node or (element.tag == word._tag("bookmarkEnd")
                                           and element.get(word._tag("id")) == ident):
                        element.getparent().remove(element)


def compare(base_data, source_data):
    """Compare arbitrary Transitional packages without assuming matching ids."""
    return Comparison(word.load(base_data, purpose="edit"), word.load(source_data, purpose="edit"))
