"""Compare and append BIFF12 cell styles without rewriting existing records."""

import copy
import math
import struct

from lxml import etree

import excel_stage_ooxml as ooxml
import excel_xlsb as binary


WorkbookError = ooxml.WorkbookError
_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_EXT_NS = "urn:excel-diff:xlsb-style-extension"
# The order is the Styles part ABNF, not the numerical record order.
_COLLECTIONS = {
    "numFmts": (615, 616, 44), "fonts": (611, 612, 43),
    "fills": (603, 604, 45), "borders": (613, 614, 46),
    "cellStyleXfs": (626, 627, 47), "cellXfs": (617, 618, 47),
    "cellStyles": (619, 620, 48),
}
_BEGIN = {value[0]: name for name, value in _COLLECTIONS.items()}
_END = {value[1]: name for name, value in _COLLECTIONS.items()}
_APPLY_BITS = (("applyNumberFormat", 0), ("applyFont", 1), ("applyFill", 4),
               ("applyBorder", 3), ("applyAlignment", 2), ("applyProtection", 5))
_PATTERNS = ("none", "solid", "mediumGray", "darkGray", "lightGray", "darkHorizontal",
             "darkVertical", "darkDown", "darkUp", "darkGrid", "darkTrellis",
             "lightHorizontal", "lightVertical", "lightDown", "lightUp", "lightGrid",
             "lightTrellis", "gray125", "gray0625")
_BORDERS = ("none", "thin", "medium", "dashed", "dotted", "thick", "double", "hair",
            "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot",
            "mediumDashDotDot", "slantDashDot")
_HORIZONTAL = ("general", "left", "center", "right", "fill", "justify", "centerContinuous", "distributed")
_VERTICAL = ("top", "center", "bottom", "justify", "distributed")
_FAILURE = "作り直したブックの書式の検証に失敗しました。書き込みは行っていません"


def _element(local, **attrs):
    return etree.Element("{" + _NS + "}" + local, {key: str(value) for key, value in attrs.items()})


def _child(parent, name, **attrs):
    node = _element(name, **attrs)
    parent.append(node)
    return node


def _extension(node, data):
    # Keep unsupported information visible to the common canonical comparison.
    child = etree.SubElement(node, "{" + _EXT_NS + "}binary")
    child.text = data.hex()


def _finish(reader, node):
    if reader.position != len(reader.data):
        _extension(node, reader.data[reader.position:])
    return node


def _decimal(value):
    if not math.isfinite(value):
        raise WorkbookError("xlsb の書式の数値が有限ではありません")
    return ooxml._dec(repr(value))


def _color(data, name="color", omit_auto=False):
    flags, index, tint, red, green, blue, alpha = struct.unpack("<BBhBBBB", data)
    kind = flags >> 1
    if kind == 0 and omit_auto:
        return None
    if kind == 0:
        node = _element(name, auto="1")
    elif kind == 1:
        node = _element(name, indexed=index)
    elif kind == 2:
        node = _element(name, rgb="%02X%02X%02X%02X" % (alpha, red, green, blue))
    elif kind == 3:
        node = _element(name, theme=index)
    else:
        node = _element(name)
        _extension(node, data)
    if tint:
        node.set("tint", _decimal(tint / 32767))
    return node


def _font(data):
    reader = binary._Reader(data)
    height, flags, weight, script, underline, family, charset, unused = reader.unpack("HHHHBBBB")
    color = _color(reader.take(8), omit_auto=True)
    scheme = reader.take(1)[0]
    name = reader.wide()
    node = _element("font")
    for flag, label in ((2, "i"), (8, "strike"), (16, "outline"), (32, "shadow"), (64, "condense"), (128, "extend")):
        if flags & flag:
            _child(node, label)
    if weight == 700:
        _child(node, "b")
    elif weight != 400:
        _extension(node, struct.pack("<H", weight))
    _child(node, "sz", val=_decimal(height / 20))
    if underline:
        _child(node, "u", val={1: "single", 2: "double", 33: "singleAccounting", 34: "doubleAccounting"}.get(underline, str(underline)))
    if script:
        _child(node, "vertAlign", val={1: "superscript", 2: "subscript"}.get(script, str(script)))
    if color is not None:
        node.append(color)
    _child(node, "name", val=name)
    if family:
        _child(node, "family", val=family)
    if charset:
        _child(node, "charset", val=charset)
    if scheme:
        _child(node, "scheme", val={1: "major", 2: "minor"}.get(scheme, str(scheme)))
    return _finish(reader, node)


def _fill(data):
    reader = binary._Reader(data)
    pattern = reader.uint()
    foreground, background = reader.take(8), reader.take(8)
    gradient_type = reader.uint()
    angle, left, right, top, bottom = reader.unpack("ddddd")
    count = reader.uint()
    if count > 256:
        raise WorkbookError("xlsb のグラデーションの停止位置の数が不正です")
    node = _element("fill")
    if pattern == 40:
        fill = _child(node, "gradientFill")
        if gradient_type == 0:
            if angle:
                fill.set("degree", _decimal(angle))
        elif gradient_type == 1:
            fill.set("type", "path")
            for name, value in (("left", left), ("right", right), ("top", top), ("bottom", bottom)):
                if value:
                    fill.set(name, _decimal(value))
        else:
            _extension(fill, struct.pack("<I", gradient_type))
        for _ in range(count):
            color = _color(reader.take(8))
            stop = _child(fill, "stop", position=_decimal(reader.unpack("d")[0]))
            stop.append(color)
    else:
        fill = _child(node, "patternFill", patternType=_PATTERNS[pattern] if pattern < len(_PATTERNS) else str(pattern))
        # Default system foreground/background are implicit in Excel's XML.
        if not (foreground[0] >> 1 == 1 and foreground[1] == 64 and foreground[2:4] == b"\0\0"):
            fill.append(_color(foreground, "fgColor"))
        if not (background[0] >> 1 == 1 and background[1] == 65 and background[2:4] == b"\0\0"):
            fill.append(_color(background, "bgColor"))
        reader.take(count * 16)
    return _finish(reader, node)


def _border(data):
    reader = binary._Reader(data)
    flags = reader.take(1)[0]
    node = _element("border")
    if flags & 1:
        node.set("diagonalDown", "1")
    if flags & 2:
        node.set("diagonalUp", "1")
    sides = {}
    for name in ("top", "bottom", "left", "right", "diagonal"):
        style, unused = reader.unpack("BB")
        color = reader.take(8)
        side = _element(name)
        if style:
            side.set("style", _BORDERS[style] if style < len(_BORDERS) else str(style))
            side.append(_color(color))
        sides[name] = side
    for name in ("left", "right", "top", "bottom", "diagonal"):
        node.append(sides[name])
    return _finish(reader, node)


def _xf(data, style=False):
    reader = binary._Reader(data)
    parent, numfmt, font, fill, border, rotation, indent, flags, apply = reader.unpack("HHHHHBBHH")
    node = _element("xf", numFmtId=numfmt, fontId=font, fillId=fill, borderId=border)
    if not style:
        node.set("xfId", str(parent))
    elif parent != 0xFFFF:
        _extension(node, struct.pack("<H", parent))
    for name, bit in _APPLY_BITS:
        flag = bool(apply & (1 << bit))
        if flag:
            node.set(name, "0" if style else "1")
    alignment = _element("alignment")
    horizontal, vertical = flags & 7, (flags >> 3) & 7
    if horizontal:
        alignment.set("horizontal", _HORIZONTAL[horizontal])
    if vertical != 2:
        alignment.set("vertical", _VERTICAL[vertical] if vertical < len(_VERTICAL) else str(vertical))
    for name, value in (("textRotation", rotation), ("indent", indent), ("readingOrder", (flags >> 10) & 3)):
        if value:
            alignment.set(name, str(value))
    for name, bit in (("wrapText", 6), ("justifyLastLine", 7), ("shrinkToFit", 8)):
        if flags & (1 << bit):
            alignment.set(name, "1")
    if alignment.attrib:
        node.append(alignment)
    locked, hidden = bool(flags & 0x1000), bool(flags & 0x2000)
    if not locked or hidden:
        protection = _child(node, "protection")
        if not locked:
            protection.set("locked", "0")
        if hidden:
            protection.set("hidden", "1")
    if flags & 0x4000:
        node.set("pivotButton", "1")
    if flags & 0x8000:
        node.set("quotePrefix", "1")
    return _finish(reader, node)


def _style(data):
    reader = binary._Reader(data)
    index, flags, builtin, level = reader.unpack("IHBB")
    node = _element("cellStyle", name=reader.wide(), xfId=index)
    if flags & 1:
        node.set("builtinId", str(builtin))
        if builtin in (1, 2):
            node.set("iLevel", str(level))
    if flags & 2:
        node.set("hidden", "1")
    if flags & 4:
        node.set("customBuiltin", "1")
    return _finish(reader, node)


def _numfmt(data):
    reader = binary._Reader(data)
    number = reader.unpack("H")[0]
    return _finish(reader, _element("numFmt", numFmtId=number, formatCode=reader.wide()))


def _decode(name, data):
    if name == "numFmts":
        return _numfmt(data)
    if name == "fonts":
        return _font(data)
    if name == "fills":
        return _fill(data)
    if name == "borders":
        return _border(data)
    if name in ("cellXfs", "cellStyleXfs"):
        return _xf(data, name == "cellStyleXfs")
    return _style(data)


class BinaryStyles(ooxml._Styles):
    """Expose the shared style comparison contract over immutable BIFF12 bytes."""

    def __init__(self, data, path=None):
        self.records = []
        self._parent_cache = {}
        self.sections, self.entries = {}, {name: [] for name in _COLLECTIONS}
        try:
            root = self._read(data) if data is not None else None
            super().__init__(etree.tostring(root) if root is not None else None, path)
            self.data = data
        except (binary.WorkbookError, struct.error, UnicodeError) as error:
            raise WorkbookError("xlsb の書式を読めません: " + str(error)) from error

    def _read(self, data):
        self.records = list(binary.records(data))
        if not self.records or self.records[0].type != 278 or self.records[-1].type != 279:
            raise WorkbookError("xlsb の styles.bin が不正です")
        root = etree.Element("{" + _NS + "}styleSheet", nsmap={None: _NS})
        current, depth, indexed = None, 0, None
        extension_start, extension_node, extension_types = None, None, []
        style_prefixes = []
        for position, record in enumerate(self.records):
            kind = record.type
            if kind in (35, 37):
                if not depth:
                    extension_start, extension_types = record.start, []
                    extension_node = root[-1][-1] if current and self.entries[current] else None
                depth += 1
            if depth:
                extension_types.append(kind)
                if kind in (36, 38):
                    depth -= 1
                    if not depth:
                        block = data[extension_start:record.end]
                        # KnownFonts is a collection-wide marker. A style AC
                        # prefix belongs to the following BrtStyle, including
                        # when it precedes the first entry in the collection.
                        known_fonts = current == "fonts" and set(extension_types) <= {37, 38, 1025}
                        style_prefix = current == "cellStyles" and (extension_types[0] == 37 or extension_node is None)
                        if style_prefix:
                            items = list(binary.records(block))
                            # The GUID identifies a style but is not part of
                            # its formatting, like cellStyle xr:uid in OOXML.
                            known_marker = (
                                extension_types == [37, 3072, 38]
                                and items[0].payload == bytes.fromhex("010002110080")
                                and len(items[1].payload) == 16
                                and not items[2].payload
                            )
                            if not known_marker:
                                style_prefixes.append(block)
                        elif not known_fonts and extension_node is not None:
                            node = _child(extension_node, "extLst") if current in ("cellXfs", "cellStyleXfs") else extension_node
                            _extension(node, block)
                continue
            if kind in (36, 38):
                raise WorkbookError("xlsb の書式の拡張の範囲が不正です")
            if kind in _BEGIN:
                name = _BEGIN[kind]
                if current is not None or name in self.sections or len(record.payload) != 4:
                    raise WorkbookError("xlsb の書式の集まりが不正です")
                current = name
                self.sections[name] = [position, None]
                root.append(_element(name, count=struct.unpack("<I", record.payload)[0]))
            elif kind in _END:
                name = _END[kind]
                if current != name:
                    raise WorkbookError("xlsb の書式の集まりの終わりが不正です")
                self.sections[name][1] = position
                if style_prefixes:
                    raise WorkbookError("xlsb のセルのスタイルの拡張の位置が不正です")
                if int(root[-1].get("count")) != len(self.entries[name]):
                    raise WorkbookError("xlsb の書式の件数が一致しません")
                current = None
            elif current is not None and kind == _COLLECTIONS[current][2]:
                self.entries[current].append(record)
                node = _decode(current, record.payload)
                if current == "cellStyles":
                    for block in style_prefixes:
                        _extension(node, block)
                    style_prefixes.clear()
                root[-1].append(node)
            elif kind == 565:
                if indexed is not None:
                    raise WorkbookError("xlsb のパレットが重複しています")
                colors = _child(root, "colors")
                indexed = _child(colors, "indexedColors")
            elif kind == 475:
                if indexed is None or len(record.payload) != 4:
                    raise WorkbookError("xlsb のパレットが不正です")
                _child(indexed, "rgbColor", rgb="FF" + record.payload[:3].hex().upper())
        if current is not None or depth:
            raise WorkbookError("xlsb の書式の集まりが閉じていません")
        formats = root.find("{" + _NS + "}numFmts")
        if formats is not None and len({node.get("numFmtId") for node in formats}) != len(self.entries["numFmts"]):
            raise WorkbookError("xlsb の表示形式の番号が重複しています")
        return root

    def parent_exact(self, index):
        if index not in self._parent_cache:
            # Reuse the common field normalization over a view of style XFs.
            # The style name is checked separately; it is not a parent field.
            view = copy.copy(self)
            view.cellXfs = self.cellStyleXfs
            fields = ooxml._Styles._fields(view, index)
            fields.pop("style", None)
            flags = tuple(ooxml._bool(self.cellStyleXfs[index].get(name), None) for name in ooxml._APPLY) if 0 <= index < len(self.cellStyleXfs) else ()
            self._parent_cache[index] = (tuple(sorted(fields.items())), flags)
        return self._parent_cache[index]

    def importer(self, source):
        return BinaryStyleImporter(self, source)


class BinaryStyleImporter(ooxml._StyleImporter):
    """Copy only required style records and append them to their collections."""

    def __init__(self, target_styles, source_styles):
        super().__init__(target_styles, source_styles)
        self.appended = {name: [] for name in _COLLECTIONS}
        self._numfmts = None
        self._styles = None

    def _needs_custom_numfmt(self, index):
        if not 0 <= index < len(self.source.cellXfs):
            return False
        xf = self.source.cellXfs[index]
        nodes = [xf]
        style = ooxml._index(xf.get("xfId"))
        if 0 <= style < len(self.source.cellStyleXfs):
            nodes.append(self.source.cellStyleXfs[style])
        existing = {self.target.numfmt(number) for number in self.target.numFmts}
        existing.update(("code", code) for number, code in ooxml._INVARIANT_NUMFMTS.items() if number not in self.target.numFmts)
        return any(self.source.numfmt(node.get("numFmtId"))[0] == "code" and self.source.numfmt(node.get("numFmtId")) not in existing for node in nodes)

    def problem(self, index):
        reason = super().problem(index)
        if reason:
            return reason
        if self.source.exact(index) in self._exact_indices():
            return None
        xf = self.source.cellXfs[index]
        source_parent = ooxml._index(xf.get("xfId"))
        identity = self.source.style_identity(xf)
        target_parent = next((number for number, other in self.target._names.items() if other == identity), None)
        if target_parent is not None:
            if self.source.parent_exact(source_parent) != self.target.parent_exact(target_parent):
                return "同じ名前で中身の違うスタイル"
        else:
            style = self.source._style_elements.get(source_parent)
            if style is not None and any(ooxml._namespace(node) != self.source.ns or any(name.startswith("{") for name in node.attrib) for node in style.iter()):
                return "書式の拡張機能"
        if self._needs_custom_numfmt(index) and any(path.startswith(("xl/pivotCache/", "xl/pivotTables/")) and path.endswith(".bin") for path in self.target.parts):
            return "ピボットの表示形式番号を確認できない"
        if self._needs_custom_numfmt(index) and any(record.type in (507, 1171, 2102) for record in self.target.records):
            return "条件付き書式の表示形式番号を確認できない"
        if self._needs_custom_numfmt(index) and self._available_numfmt() > 65535:
            return "表示形式の番号が上限を超える"
        return None

    def _append_record(self, name, payload):
        count = len(self.target.entries[name]) + len(self.appended[name])
        if name in ("cellXfs", "cellStyleXfs") and count >= 64000:
            raise WorkbookError("セルの書式またはスタイルの数が 64000 を超えます")
        if name in ("fonts", "fills", "borders") and count >= 65536:
            raise WorkbookError("書式の参照番号が上限を超えます")
        self.appended[name].append(payload)
        self.modified = True
        return count

    def _available_numfmt(self):
        numbers = [163]
        numbers.extend(self.target.numFmts)
        for name in ("cellXfs", "cellStyleXfs"):
            numbers.extend(ooxml._index(node.get("numFmtId")) for node in getattr(self.target, name))
        for path, data in self.target.parts.items():
            if path.startswith(("xl/pivotCache/", "xl/pivotTables/")) and path.endswith(".xml"):
                tree = ooxml._xml(data, path)
                numbers.extend(ooxml._index(value, -1) for node in tree.iter() for attr, value in node.attrib.items() if ooxml._local_attribute(attr) == "numFmtId")
        return max(numbers) + 1

    def _numfmt(self, value):
        fmt = self.source.numfmt(value)
        if fmt[0] == "builtin":
            return fmt[1]
        if self._numfmts is None:
            self._numfmts = {node.get("formatCode", ""): number for number, node in self.target.numFmts.items()}
            for number, code in ooxml._INVARIANT_NUMFMTS.items():
                if number not in self.target.numFmts:
                    self._numfmts.setdefault(code, number)
        code = fmt[1]
        if code in self._numfmts:
            return self._numfmts[code]
        if self._next_numfmt is None:
            self._next_numfmt = self._available_numfmt()
        number = self._next_numfmt
        if number > 65535:
            raise WorkbookError("書式は反映できません: 表示形式の番号が上限を超える")
        self._next_numfmt += 1
        encoded = code.encode("utf-16le")
        self._append_record("numFmts", struct.pack("<HI", number, len(encoded) // 2) + encoded)
        self._numfmts[code] = number
        return number

    def _component(self, name, number, flatten):
        if name not in self._components:
            mapping = {}
            for index, node in enumerate(getattr(self.target, name)):
                mapping.setdefault(tuple(sorted(flatten(node).items())), index)
            self._components[name] = mapping
        node = getattr(self.source, name)[number]
        key = tuple(sorted(flatten(node).items()))
        mapping = self._components[name]
        if key not in mapping:
            mapping[key] = self._append_record(name, self.source.entries[name][number].payload)
        return mapping[key]

    def _remap(self, name, number):
        original = getattr(self.source, name)[number]
        data = bytearray(self.source.entries[name][number].payload)
        for offset, value in (
            (2, self._numfmt(original.get("numFmtId"))),
            (4, self._component("fonts", ooxml._index(original.get("fontId")), ooxml._font_fields)),
            (6, self._component("fills", ooxml._index(original.get("fillId")), ooxml._fill_fields)),
            (8, self._component("borders", ooxml._index(original.get("borderId")), ooxml._border_fields)),
        ):
            struct.pack_into("<H", data, offset, value)
        return data

    def _style(self, xf):
        if self._styles is None:
            self._styles = {identity: index for index, identity in self.target._names.items()}
        identity = self.source.style_identity(xf)
        if identity in self._styles:
            return self._styles[identity]
        source_index = ooxml._index(xf.get("xfId"))
        if identity[0] == "unnamed" and not self.source.cellStyleXfs and not self.target.cellStyleXfs:
            return 0
        payload = self._remap("cellStyleXfs", source_index)
        number = self._append_record("cellStyleXfs", bytes(payload))
        if identity[0] != "unnamed":
            node = self.source._style_elements[source_index]
            style_index = next(index for index, candidate in enumerate(self.source.cellStyles) if candidate is node)
            payload = bytearray(self.source.entries["cellStyles"][style_index].payload)
            struct.pack_into("<I", payload, 0, number)
            self._append_record("cellStyles", bytes(payload))
        self._styles[identity] = number
        return number

    def index_for(self, index):
        if index in self._indices:
            return self._indices[index]
        reason = self.problem(index)
        if reason:
            raise WorkbookError("書式は反映できません: " + reason)
        if self.target.data == self.source.data:
            result = index
        elif self.source.exact(index) in self._exact_indices():
            result = self._exact_indices()[self.source.exact(index)]
        else:
            node = self.source.cellXfs[index]
            payload = self._remap("cellXfs", index)
            struct.pack_into("<H", payload, 0, self._style(node))
            result = self._append_record("cellXfs", bytes(payload))
            self._exact_indices().setdefault(self.source.exact(index), result)
        self._indices[index] = result
        return result

    def serialize(self, original_bytes):
        if not self.modified:
            return original_bytes
        before, after = {}, {}
        missing = [name for name in _COLLECTIONS if self.appended[name] and name not in self.target.sections]
        for name in missing:
            if name not in ("numFmts", "cellStyles"):
                raise WorkbookError("書式は反映できません: 取り込み先に書式の定義がない")
            following = next((other for other in _COLLECTIONS if list(_COLLECTIONS).index(other) > list(_COLLECTIONS).index(name) and other in self.target.sections), None)
            position = self.target.sections[following][0] if following else len(self.target.records) - 1
            begin, end, kind = _COLLECTIONS[name]
            payloads = self.appended[name]
            content = _record(begin, struct.pack("<I", len(payloads))) + b"".join(_record(kind, payload) for payload in payloads) + _record(end, b"")
            before.setdefault(position, []).append(content)
        for name, (first, last) in self.target.sections.items():
            payloads = self.appended[name]
            if payloads:
                begin, end, kind = _COLLECTIONS[name]
                after[first] = _record(begin, struct.pack("<I", len(self.target.entries[name]) + len(payloads)))
                before.setdefault(last, []).append(b"".join(_record(kind, payload) for payload in payloads))
        output = []
        for position, record in enumerate(self.target.records):
            output.extend(before.get(position, ()))
            output.append(after.get(position, original_bytes[record.start:record.end]))
        return b"".join(output)


def _record(kind, payload):
    return binary.record_header(kind, len(payload)) + payload


def check_styles_appended(original_bytes, result_bytes):
    """Verify original bytes, permitting only collection counts and tail entries."""
    try:
        original, result = BinaryStyles(original_bytes), BinaryStyles(result_bytes)
        if not set(original.sections) <= set(result.sections) or set(result.sections) - set(original.sections) - {"numFmts", "cellStyles"}:
            raise WorkbookError(_FAILURE)
        # Remove new entries and restore old counts. The entire original stream,
        # including unknown AC/FRT records and their original headers, must match.
        remove, replace = set(), {}
        for name, (first, last) in result.sections.items():
            entries = result.entries[name]
            if name not in original.sections:
                begin, end, kind = _COLLECTIONS[name]
                if [record.type for record in result.records[first:last + 1]] != [begin] + [kind] * len(entries) + [end]:
                    raise WorkbookError(_FAILURE)
                remove.update(range(first, last + 1))
                continue
            old = original.entries[name]
            if len(entries) < len(old):
                raise WorkbookError(_FAILURE)
            for left, right in zip(old, entries):
                if original_bytes[left.start:left.end] != result_bytes[right.start:right.end]:
                    raise WorkbookError(_FAILURE)
            if len(entries) > len(old):
                tail = entries[len(old):]
                # Appends must be one contiguous suffix directly before the end.
                first_tail = next(index for index, record in enumerate(result.records) if record is tail[0])
                if result.records[first_tail:last] != tail:
                    raise WorkbookError(_FAILURE)
                remove.update(range(first_tail, last))
            original_begin = original.records[original.sections[name][0]]
            result_begin = result.records[first]
            if original_bytes[original_begin.start:original_begin.payload_start] != result_bytes[result_begin.start:result_begin.payload_start]:
                raise WorkbookError(_FAILURE)
            replace[first] = original_bytes[original_begin.start:original_begin.end]
        restored = b"".join(replace.get(index, result_bytes[record.start:record.end]) for index, record in enumerate(result.records) if index not in remove)
        if restored != original_bytes:
            raise WorkbookError(_FAILURE)
        ordered = [list(_COLLECTIONS).index(name) for name in sorted(result.sections, key=lambda name: result.sections[name][0])]
        if ordered != sorted(ordered):
            raise WorkbookError(_FAILURE)
    except (WorkbookError, binary.WorkbookError, ValueError, TypeError) as error:
        raise WorkbookError(_FAILURE) from error


_BinaryStyles = BinaryStyles
_BinaryStyleImporter = BinaryStyleImporter
_check_styles_appended = check_styles_appended
