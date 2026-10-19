"""Lossless package editing for staging ordinary OOXML cell values/formulas.

Only the selected cells are reconstructed.  In particular, this module does
not save a workbook through openpyxl, which can discard unsupported features.
"""

import copy
import hashlib
import io
import json
import posixpath
import re
import zipfile
from array import array
from bisect import bisect_right
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from json.encoder import encode_basestring

from lxml import etree


class WorkbookError(ValueError):
    """A workbook or requested edit cannot safely be partially staged."""


@dataclass(frozen=True)
class Cell:
    __slots__ = ("kind", "value")

    kind: str
    value: object


@dataclass(frozen=True)
class Sheet:
    id: str
    name: str
    path: str


_COORDINATE = re.compile(r"^([A-Z]{1,3})([1-9][0-9]{0,6})$")
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_KINDS = {"number", "string", "boolean", "error", "date", "formula"}
FORMAT_KIND = "format"
_RECORD_KINDS = _KINDS | {FORMAT_KIND}
_NUMBER_TEXT = re.compile(r"([+-]?)([0-9]*)(?:\.([0-9]*))?(?:[eE]([+-]?[0-9]{1,6}))?\Z")
_JSON_KINDS = {kind: encode_basestring(kind) for kind in _RECORD_KINDS}
# Japanese Excel saves readings (furigana) beside plain text: rPh holds the
# reading and phoneticPr its settings. Neither is rich text formatting.
_PLAIN_STRING = {"t", "rPh", "phoneticPr"}


def coordinate_key(coordinate):
    match = _COORDINATE.fullmatch(coordinate)
    if not match:
        raise WorkbookError("不正なセル番地です: " + str(coordinate))
    column = 0
    for letter in match[1]:
        column = column * 26 + ord(letter) - ord("A") + 1
    row = int(match[2])
    if column > 16384 or row > 1048576:
        raise WorkbookError("Excel の範囲外のセル番地です: " + coordinate)
    return row, column


def _decimal_number(value):
    try:
        number = Decimal(value)
    except (InvalidOperation, ValueError):
        raise WorkbookError("不正な数値セルです: " + str(value)) from None
    if not number.is_finite():
        raise WorkbookError("有限でない数値セルは扱えません")
    if number.is_zero():
        return "0"
    # Decimal.normalize() rounds to its context precision.  Remove trailing
    # zeros from the tuple instead, retaining every digit from the workbook.
    sign, digits, exponent = number.as_tuple()
    digits = list(digits)
    while digits[-1] == 0:
        digits.pop()
        exponent += 1
    canonical = Decimal((sign, tuple(digits), exponent))
    return format(canonical, "f") if -6 <= canonical.adjusted() < 16 else str(canonical)


def _number(value):
    if isinstance(value, str):
        # Most worksheet numbers are already ordinary, positive integers.
        if value.isascii() and value.isdecimal():
            digits = value.lstrip("0")
            if len(digits) <= 16:
                return digits or "0"
        # Keep Decimal as the authority for its additional accepted spellings
        # (whitespace, underscores, Unicode digits, very large exponents).
        match = _NUMBER_TEXT.fullmatch(value)
        if match and (match[2] or match[3]):
            fraction = match[3] or ""
            digits = (match[2] + fraction).lstrip("0")
            if not digits:
                return "0"
            exponent = int(match[4] or "0") - len(fraction)
            significant = digits.rstrip("0")
            exponent += len(digits) - len(significant)
            adjusted = len(significant) + exponent - 1
            sign = "-" if match[1] == "-" else ""
            # str(Decimal) also keeps fixed notation for nonpositive
            # exponents, even when adjusted is 16 or larger.
            if adjusted >= -6 and (adjusted < 16 or exponent <= 0):
                point = len(significant) + exponent
                if exponent >= 0:
                    result = significant + "0" * exponent
                elif point > 0:
                    result = significant[:point] + "." + significant[point:]
                else:
                    result = "0." + "0" * -point + significant
            else:
                result = significant[0]
                if len(significant) > 1:
                    result += "." + significant[1:]
                result += "E" + ("+" if adjusted >= 0 else "") + str(adjusted)
            return sign + result
    return _decimal_number(value)


def _cell_record(coordinate, cell):
    # Coordinates have already been checked and contain ASCII letters/digits.
    prefix = '["' + coordinate + '",'
    if cell is None:
        return prefix + '"blank",null]'
    kind, value = cell.kind, cell.value
    if type(kind) is str:
        encoded_kind = _JSON_KINDS.get(kind)
        if encoded_kind is None:
            encoded_kind = encode_basestring(kind)
        if isinstance(value, str):
            encoded_value = encode_basestring(value)
        elif value is True:
            encoded_value = "true"
        elif value is False:
            encoded_value = "false"
        elif value is None:
            encoded_value = "null"
        else:
            encoded_value = None
        if encoded_value is not None:
            return prefix + encoded_kind + "," + encoded_value + "]"
    return json.dumps([coordinate, kind, value], ensure_ascii=False, allow_nan=False, separators=(",", ":"))


def cell_record(coordinate, cell):
    coordinate_key(coordinate)
    return _cell_record(coordinate, cell)


def parse_record(line):
    try:
        values = json.loads(line)
    except (ValueError, TypeError):
        raise WorkbookError("セル選択の行が不正です") from None
    if not isinstance(values, list) or len(values) != 3 or not isinstance(values[0], str):
        raise WorkbookError("セル選択の行が不正です")
    coordinate, kind, value = values
    coordinate_key(coordinate)
    if not isinstance(kind, str):
        raise WorkbookError("セル選択の型が不正です")
    if kind == "blank" and value is None:
        return coordinate, None
    if kind not in _RECORD_KINDS:
        raise WorkbookError("未知のセル型です: " + str(kind))
    if kind == "boolean":
        if not isinstance(value, bool):
            raise WorkbookError("真偽値セルの選択が不正です")
    elif not isinstance(value, str):
        raise WorkbookError("セル選択の値が不正です")
    if kind == "number" and _number(value) != value:
        raise WorkbookError("数値セルの選択が正規化されていません")
    if kind == "formula" and not value.startswith("="):
        raise WorkbookError("数式セルの選択が不正です")
    return coordinate, Cell(kind, value)


def _local(node):
    # Read Clark-notation tags ("{namespace}name") directly; creating a QName
    # for every node dominated the time spent on large worksheets.
    tag = node.tag
    return tag.rpartition("}")[2] if isinstance(tag, str) else ""


def _namespace(node):
    tag = node.tag
    return tag[1:tag.index("}")] if isinstance(tag, str) and tag.startswith("{") else ""


def _tag(namespace, name):
    return "{" + namespace + "}" + name if namespace else name


def _children(node, name):
    tag = _tag(_namespace(node), name)
    return [child for child in node if child.tag == tag]


def _child(node, name):
    children = _children(node, name)
    if len(children) > 1:
        raise WorkbookError("XML の要素が重複しています: " + name)
    return children[0] if children else None


def _xml(data, description):
    try:
        parser = etree.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False)
        tree = etree.parse(io.BytesIO(data), parser)
        if tree.docinfo.doctype:
            raise WorkbookError("DTD を含む XML は扱えません: " + description)
        return tree
    except (etree.XMLSyntaxError, ValueError) as error:
        raise WorkbookError("Excel の XML を読めません: " + description) from error


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


def _target(owner, target):
    path = posixpath.normpath(posixpath.join(posixpath.dirname(owner), target))
    if target.startswith("/"):
        path = posixpath.normpath(target.lstrip("/"))
    if path in ("", ".", "..") or path.startswith("../") or "\\" in path:
        raise WorkbookError("不正な Excel パーツの参照です")
    return path


def _rels_path(owner):
    return posixpath.join(posixpath.dirname(owner), "_rels", posixpath.basename(owner) + ".rels")


def _range(reference):
    ends = reference.replace("$", "").split(":")
    if len(ends) == 1:
        ends *= 2
    if len(ends) != 2:
        raise WorkbookError("不正なセル範囲です: " + reference)
    first, last = [coordinate_key(value) for value in ends]
    if first[0] > last[0] or first[1] > last[1]:
        raise WorkbookError("不正なセル範囲です: " + reference)
    return first, last


def _inside(coordinate, rectangle):
    row, column = coordinate_key(coordinate)
    first, last = rectangle
    return first[0] <= row <= last[0] and first[1] <= column <= last[1]


def _string(node):
    # Phonetic text (rPh) is presentation, not the cell string.
    return "".join(child.text or "" for child in node.iter() if _local(child) == "t" and _local(child.getparent()) != "rPh")


_INVARIANT_NUMFMTS = {
    0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00",
    9: "0%", 10: "0.00%", 11: "0.00E+00", 12: "# ?/?", 13: "# ??/??",
    20: "h:mm", 21: "h:mm:ss", 45: "mm:ss", 46: "[h]:mm:ss",
    47: "mmss.0", 48: "##0.0E+0", 49: "@",
}
_DEFAULT_PALETTE = tuple((
    "000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF "
    "000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF "
    "800000 008000 000080 808000 800080 008080 C0C0C0 808080 "
    "9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF "
    "000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF "
    "00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 "
    "3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 "
    "003366 339966 003300 333300 993300 993366 333399 333333"
).split())
_STYLE_CONTAINER_ORDER = (
    "numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs",
    "cellStyles", "dxfs", "tableStyles", "colors", "extLst",
)
_APPLY = ("applyNumberFormat", "applyFont", "applyFill", "applyBorder", "applyAlignment", "applyProtection")


def _bool(value, default=False):
    if value is None:
        return default
    if value in ("1", "true"):
        return True
    if value in ("0", "false"):
        return False
    return ("raw", value)


def _dec(value):
    try:
        return _number(value)
    except (WorkbookError, TypeError):
        return value


def _index(value, default=0):
    if value is None:
        return default
    try:
        result = int(value)
        return result if result >= 0 else -1
    except (TypeError, ValueError):
        return -1


def _raw(element):
    return etree.tostring(element, method="c14n", exclusive=True, with_comments=False).decode("utf-8")


def _color(element):
    if element is None:
        return None
    known = {"rgb", "theme", "indexed", "auto", "tint"}
    if any(name not in known for name in element.attrib) or len(element):
        return ("raw", _raw(element))
    if element.get("rgb") is not None:
        value = ("rgb", element.get("rgb").upper()[-6:])
    elif element.get("theme") is not None:
        value = ("theme", _dec(element.get("theme")))
    elif element.get("indexed") is not None:
        value = ("indexed", _dec(element.get("indexed")))
    elif _bool(element.get("auto")):
        value = ("auto",)
    else:
        return None
    tint = _dec(element.get("tint", "0"))
    return value + (("tint", tint),) if tint != "0" else value


def _font_fields(element):
    fields = {}
    for node in element:
        name = _local(node)
        if _namespace(node) != _namespace(element) or (name != "color" and (set(node.attrib) - {"val"} or len(node))):
            fields["raw." + node.tag] = _raw(node)
        elif name in {"b", "i", "strike", "condense", "extend", "outline", "shadow"}:
            value = _bool(node.get("val"), True)
            if value is not False:
                fields[name] = value
        elif name == "color":
            value = _color(node)
            if value is not None:
                fields[name] = value
        elif name == "u":
            value = node.get("val", "single")
            if value != "none":
                fields[name] = value
        elif name == "vertAlign":
            if node.get("val", "baseline") != "baseline":
                fields[name] = node.get("val")
        elif name == "scheme":
            if node.get("val", "none") != "none":
                fields[name] = node.get("val")
        elif name == "sz":
            fields[name] = _dec(node.get("val"))
        elif name in {"name", "family", "charset"}:
            fields[name] = node.get("val")
        else:
            fields["raw." + name] = _raw(node)
    for name, value in element.attrib.items():
        fields["attr." + name] = value
    return fields


def _fill_fields(element):
    if not len(element) and not element.attrib:
        return {}
    if len(element) == 1 and _namespace(element[0]) == _namespace(element):
        node = element[0]
        if _local(node) == "patternFill":
            pattern = node.get("patternType", "none")
            fg, bg = _color(_child(node, "fgColor")), _color(_child(node, "bgColor"))
            if set(node.attrib) <= {"patternType"} and all(_local(c) in {"fgColor", "bgColor"} and _namespace(c) == _namespace(node) for c in node):
                result = {}
                if pattern != "none" or (node.get("patternType") is None and (fg is not None or bg is not None)):
                    result["pattern"] = node.get("patternType", "(未指定)")
                if fg is not None:
                    result["fg"] = fg
                if bg is not None and pattern != "solid":
                    result["bg"] = bg
                return result
        elif _local(node) == "gradientFill":
            return {"gradient": _raw(node)}
    return {"raw": _raw(element)}


def _border_fields(element):
    fields = {}
    up, down = _bool(element.get("diagonalUp")), _bool(element.get("diagonalDown"))
    outline = _bool(element.get("outline"), True)
    if outline is not True:
        fields["outline"] = outline
    for name, value in element.attrib.items():
        if name not in {"diagonalUp", "diagonalDown", "outline"}:
            fields["attr." + name] = value
    for node in element:
        name = {"start": "left", "end": "right"}.get(_local(node), _local(node))
        if _namespace(node) != _namespace(element) or name not in {
            "left", "right", "top", "bottom", "diagonal", "vertical", "horizontal",
        } or set(node.attrib) - {"style"} or any(_local(c) != "color" or _namespace(c) != _namespace(node) for c in node):
            fields["raw." + node.tag] = _raw(node)
            continue
        style = node.get("style", "none")
        if style != "none" and (name != "diagonal" or up or down):
            value = (style, _color(_child(node, "color")))
            fields[name] = value + (up, down) if name == "diagonal" else value
    return fields


_ALIGN_DEFAULTS = {
    "horizontal": "general", "vertical": "bottom", "textRotation": "0",
    "wrapText": False, "indent": "0", "relativeIndent": "0",
    "justifyLastLine": False, "shrinkToFit": False, "readingOrder": "0",
}


def _alignment_fields(element):
    fields = {}
    for name, value in element.attrib.items():
        default = _ALIGN_DEFAULTS.get(name)
        value = _bool(value) if isinstance(default, bool) else (_dec(value) if default == "0" else value)
        if value != default:
            fields[name] = value
    for node in element:
        fields["raw." + node.tag] = _raw(node)
    return fields


def _protection_fields(element):
    fields = {}
    for name, value in element.attrib.items():
        default = True if name == "locked" else False if name == "hidden" else None
        value = _bool(value, default) if default is not None else value
        if value != default:
            fields[name] = value
    for node in element:
        fields["raw." + node.tag] = _raw(node)
    return fields


class _Styles:
    def __init__(self, data, path):
        self.data, self.path = data, path
        self.tree = _xml(data, path or "styles.xml") if data is not None else None
        self.root = self.tree.getroot() if self.tree is not None else None
        self.ns = _namespace(self.root) if self.root is not None else ""
        self._cache, self._exact_cache = {}, {}
        self.parts = {}
        self.containers = {}
        if self.root is not None:
            if _local(self.root) != "styleSheet":
                raise WorkbookError("Excel の styles XML が不正です")
            for node in self.root:
                name = _local(node)
                if name in self.containers:
                    raise WorkbookError("XML の要素が重複しています: " + name)
                self.containers[name] = node
        for name in ("fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles"):
            container = self.containers.get(name)
            setattr(self, name, list(container) if container is not None else [])
        container = self.containers.get("numFmts")
        self.numFmts = {_index(node.get("numFmtId"), -1): node for node in container} if container is not None else {}
        self._names, self._style_elements = {}, {}
        for node in self.cellStyles:
            index = _index(node.get("xfId"), -1)
            if index not in self._names:
                builtin = node.get("builtinId")
                identity = ("builtin", _dec(builtin), _dec(node.get("iLevel", "0")) if builtin in ("1", "2") else "0") if builtin is not None else ("custom", node.get("name", "").casefold())
                self._names[index] = identity
                self._style_elements[index] = node
        colors = self.containers.get("colors")
        indexed = _child(colors, "indexedColors") if colors is not None else None
        self.palette = tuple(node.get("rgb", "").upper()[-6:] for node in indexed) if indexed is not None else _DEFAULT_PALETTE

    def numfmt(self, value):
        index = _index(value)
        if index < 0:
            return ("invalid", str(value))
        if index in self.numFmts:
            return ("code", self.numFmts[index].get("formatCode", ""))
        if index in _INVARIANT_NUMFMTS:
            return ("code", _INVARIANT_NUMFMTS[index])
        return ("builtin", index) if index < 164 else ("undefined", index)

    def style_identity(self, xf):
        index = _index(xf.get("xfId"))
        if index < 0 or (index >= len(self.cellStyleXfs) and (xf.get("xfId") is not None or self.cellStyleXfs)):
            return ("invalid", str(xf.get("xfId", "0")))
        return self._names.get(index, ("unnamed",))

    def style_names(self):
        return {identity: self._style_elements[index].get("name", "") for index, identity in self._names.items()}

    def fields(self, index):
        if index not in self._cache:
            self._cache[index] = self._fields(index)
        return self._cache[index]

    def _fields(self, index):
        if self.data is None and index == 0:
            return {}
        if not isinstance(index, int) or index < 0 or index >= len(self.cellXfs):
            return {"invalid": str(index)}
        xf = self.cellXfs[index]
        fields = {}
        fmt = self.numfmt(xf.get("numFmtId"))
        if fmt != ("code", "General"):
            fields["numFmt"] = fmt
        for attr, name, flatten, prefix in (
            ("fontId", "fonts", _font_fields, "font."),
            ("fillId", "fills", _fill_fields, "fill."),
            ("borderId", "borders", _border_fields, "border."),
        ):
            number, entries = _index(xf.get(attr)), getattr(self, name)
            if number < 0 or number >= len(entries):
                fields["invalid." + attr] = str(xf.get(attr, "0"))
            else:
                fields.update((prefix + k, v) for k, v in flatten(entries[number]).items())
        fields["style"] = self.style_identity(xf)
        for name in ("quotePrefix", "pivotButton"):
            value = _bool(xf.get(name))
            if value is not False:
                fields[name] = value
        for name, value in xf.attrib.items():
            if name not in {"numFmtId", "fontId", "fillId", "borderId", "xfId", "quotePrefix", "pivotButton", *_APPLY}:
                fields["attr." + name] = value
        for node in xf:
            name = _local(node)
            if _namespace(node) == self.ns and name in ("alignment", "protection"):
                prefix, flatten = ("align.", _alignment_fields) if name == "alignment" else ("prot.", _protection_fields)
                fields.update((prefix + k, v) for k, v in flatten(node).items())
            else:
                fields["raw." + node.tag] = _raw(node)
        return fields

    def canonical(self, index):
        return tuple(sorted(self.fields(index).items()))

    def exact(self, index):
        if index not in self._exact_cache:
            flags = tuple(_bool(self.cellXfs[index].get(name), None) for name in _APPLY) if 0 <= index < len(self.cellXfs) else ()
            self._exact_cache[index] = (self.canonical(index), flags)
        return self._exact_cache[index]

    def problem(self, index):
        if index < 0 or index >= len(self.cellXfs):
            return "セルの書式番号が不正"
        xf = self.cellXfs[index]
        if any(_local(node) == "extLst" for node in xf):
            return "書式の拡張機能"
        fmt = self.numfmt(xf.get("numFmtId"))
        if fmt[0] in ("undefined", "invalid"):
            return "表示形式の定義がないか番号が不正"
        elements = []
        for attr, name in (("fontId", "fonts"), ("fillId", "fills"), ("borderId", "borders")):
            number = _index(xf.get(attr))
            if number < 0 or number >= len(getattr(self, name)):
                return "書式の参照番号が不正"
            elements.append(getattr(self, name)[number])
        number = _index(xf.get("xfId"))
        if self.style_identity(xf)[0] == "invalid":
            return "セルのスタイル番号が不正"
        fmt_node = self.numFmts.get(_index(xf.get("numFmtId")))
        if fmt_node is not None:
            elements.append(fmt_node)
        if number < len(self.cellStyleXfs):
            style_xf = self.cellStyleXfs[number]
            if any(_local(node) == "extLst" for node in style_xf):
                return "書式の拡張機能"
            elements.append(style_xf)
            for attr, name in (("fontId", "fonts"), ("fillId", "fills"), ("borderId", "borders")):
                component = _index(style_xf.get(attr))
                if component < 0 or component >= len(getattr(self, name)):
                    return "書式の参照番号が不正"
                elements.append(getattr(self, name)[component])
            if self.numfmt(style_xf.get("numFmtId"))[0] in ("undefined", "invalid"):
                return "表示形式の定義がないか番号が不正"
            fmt_node = self.numFmts.get(_index(style_xf.get("numFmtId")))
            if fmt_node is not None:
                elements.append(fmt_node)
        if any(_namespace(child) != self.ns or any(name.startswith("{") for name in child.attrib) for element in elements for child in element.iter() if isinstance(child.tag, str)):
            return "書式の拡張名前空間"
        return None

    def uses_low_indexed(self, index):
        if not 0 <= index < len(self.cellXfs):
            return False
        xf = self.cellXfs[index]
        nodes = [xf]
        style = _index(xf.get("xfId"))
        if 0 <= style < len(self.cellStyleXfs):
            nodes.append(self.cellStyleXfs[style])
        for node in list(nodes):
            for attr, name in (("fontId", "fonts"), ("fillId", "fills"), ("borderId", "borders")):
                number = _index(node.get(attr))
                if 0 <= number < len(getattr(self, name)):
                    nodes.append(getattr(self, name)[number])
        return any(0 <= _index(child.get("indexed"), -1) < 64 for node in nodes for child in node.iter())


_LABELS = {
    "numFmt": "表示形式", "font.name": "フォント", "font.sz": "サイズ", "font.b": "太字", "font.i": "斜体",
    "font.u": "下線", "font.strike": "取り消し線", "font.vertAlign": "上付き・下付き", "font.color": "文字色",
    "font.scheme": "テーマのフォント", "font.family": "フォントの系統", "font.charset": "文字セット",
    "font.outline": "輪郭", "font.shadow": "影", "font.condense": "文字幅の縮小", "font.extend": "文字幅の拡張",
    "fill.pattern": "塗りつぶし", "fill.fg": "塗りつぶしの色", "fill.bg": "パターンの背景色", "fill.gradient": "グラデーション",
    "border.left": "罫線(左)", "border.right": "罫線(右)", "border.top": "罫線(上)", "border.bottom": "罫線(下)",
    "border.diagonal": "罫線(斜め)", "border.vertical": "罫線(内側の縦)", "border.horizontal": "罫線(内側の横)", "border.outline": "罫線の外枠",
    "align.horizontal": "横位置", "align.vertical": "縦位置", "align.wrapText": "折り返し",
    "align.shrinkToFit": "縮小", "align.indent": "インデント", "align.textRotation": "文字の向き",
    "align.readingOrder": "文字の方向", "align.justifyLastLine": "前後のスペース", "align.relativeIndent": "相対インデント",
    "prot.locked": "ロック", "prot.hidden": "数式の表示", "quotePrefix": "文字列の接頭辞", "pivotButton": "ピボットのボタン", "style": "スタイル",
}
_BOOLEAN_TEXT = {
    **{"font." + key: (name, name + "なし") for key, name in (("b", "太字"), ("i", "斜体"), ("strike", "取り消し線"), ("outline", "輪郭"), ("shadow", "影"))},
    "font.condense": ("文字幅の縮小", "文字幅の縮小なし"), "font.extend": ("文字幅の拡張", "文字幅の拡張なし"),
    "align.wrapText": ("折り返して全体を表示", "折り返しなし"), "align.shrinkToFit": ("縮小して全体を表示", "縮小なし"),
    "align.justifyLastLine": ("前後にスペースを入れる", "前後のスペースなし"),
    "prot.locked": ("ロック", "ロックなし"), "prot.hidden": ("数式を表示しない", "数式を表示する"),
    "quotePrefix": ("文字列として入力(')", "文字列の接頭辞なし"), "pivotButton": ("ピボットのボタン", "ピボットのボタンなし"),
    "border.outline": ("罫線の外枠あり", "罫線の外枠なし"),
}
_STYLE_VALUES = {
    "single": "一重線", "double": "二重線", "singleAccounting": "一重線(会計)", "doubleAccounting": "二重線(会計)",
    "superscript": "上付き", "subscript": "下付き", "minor": "本文", "major": "見出し",
    "solid": "単色", "gray125": "12.5% 灰色", "gray0625": "6.25% 灰色", "lightGray": "25% 灰色", "mediumGray": "50% 灰色", "darkGray": "75% 灰色",
    "thin": "細線", "hair": "極細線", "medium": "中線", "thick": "太線", "dotted": "点線", "dashed": "破線",
    "dashDot": "一点鎖線", "dashDotDot": "二点鎖線", "mediumDashed": "中破線", "mediumDashDot": "中一点鎖線",
    "mediumDashDotDot": "中二点鎖線", "slantDashDot": "斜め一点鎖線",
}
_HORIZONTAL = {"general": "標準", "left": "左詰め", "center": "中央揃え", "right": "右詰め", "fill": "繰り返し", "justify": "両端揃え", "centerContinuous": "選択範囲内で中央", "distributed": "均等割り付け"}
_VERTICAL = {"top": "上詰め", "center": "中央揃え", "bottom": "下詰め", "justify": "両端揃え", "distributed": "均等割り付け"}


def _field_default(key):
    if key == "numFmt":
        return ("code", "General")
    if key in _BOOLEAN_TEXT:
        return key in ("prot.locked", "border.outline")
    if key.startswith("align."):
        return _ALIGN_DEFAULTS.get(key[6:])
    return None


def _describe_color(value):
    if value is None:
        return "既定"
    kind = value[0]
    if kind == "rgb":
        text = "#" + value[1]
    elif kind == "theme":
        text = "テーマ" + str(value[1])
    elif kind == "indexed":
        text = "パレット" + str(value[1])
    elif kind == "auto":
        text = "自動"
    else:
        return str(value)
    if value[-1] and isinstance(value[-1], tuple) and value[-1][0] == "tint":
        text += "(濃淡" + str(value[-1][1]) + ")"
    return text


def _describe_value(key, value, names):
    if key == "numFmt":
        kind, code = value
        return "G/標準" if value == ("code", "General") else ("「" + code + "」" if kind == "code" else ("組み込み" if kind == "builtin" else "未定義") + str(code))
    if key == "style":
        return "(名前なし)" if value is None or value[0] == "unnamed" else "「" + names.get(value, str(value)) + "」"
    if value is None:
        if key == "fill.pattern" or key.startswith("border.") and key != "border.outline":
            return "なし"
        return "既定"
    if key in {"font.color", "fill.fg", "fill.bg"}:
        return _describe_color(value)
    if key.startswith("border.") and isinstance(value, tuple) and len(value) in (2, 4):
        text = _STYLE_VALUES.get(value[0], value[0])
        if value[1] is not None:
            text += " " + _describe_color(value[1])
        if len(value) == 4:
            text += " " + ("両方" if value[2] and value[3] else "右上がり" if value[2] else "右下がり")
        return text
    if key == "align.horizontal":
        return _HORIZONTAL.get(value, str(value))
    if key == "align.vertical":
        return _VERTICAL.get(value, str(value))
    if key == "align.readingOrder":
        return {"0": "コンテキスト", "1": "左から右", "2": "右から左"}.get(value, str(value))
    if key == "align.textRotation":
        try:
            number = int(value)
            return "縦書き" if number == 255 else str(number if number <= 90 else -(number - 90)) + "度"
        except (ValueError, TypeError):
            return str(value)
    if key == "font.name":
        return "「" + str(value) + "」"
    return _STYLE_VALUES.get(value, str(value)) if isinstance(value, str) else str(value)


def _description(fields, reference, names):
    output = []
    ordered = list(_LABELS) + sorted(set(fields) | set(reference) - set(_LABELS))
    for key in dict.fromkeys(ordered):
        default = _field_default(key)
        value = fields.get(key, default)
        if value == reference.get(key, default):
            continue
        if key in _BOOLEAN_TEXT and isinstance(value, bool):
            output.append(_BOOLEAN_TEXT[key][0 if value else 1])
        else:
            label = _LABELS.get(key, "その他(" + key + ")")
            output.append(label + "=" + _describe_value(key, value, names))
    text = "、".join(output) or "既定の書式"
    return re.sub(r"[\x00-\x1f\x7f\x85\u2028\u2029]", lambda match: "\\u%04x" % ord(match[0]), text)


class FormatComparison:
    def __init__(self, target, source):
        target.assert_compatible(source)
        self.target, self.source = target, source
        self.target_styles = target.stylesheet()
        other = source.stylesheet()
        self.source_styles = self.target_styles if other.data == self.target_styles.data else other
        self._ids, self._target_ids, self._source_ids = {}, {}, {}
        self._texts, self._problems = {}, {}
        self._display_names = (self.target_styles.style_names(), self.source_styles.style_names())
        self.reference = self.target_styles.fields(0) if not self.target_styles.problem(0) else {}
        self._importer = None

    def _identifier(self, styles, index, cache):
        if index not in cache:
            canonical = styles.canonical(index)
            cache[index] = self._ids.setdefault(canonical, len(self._ids))
        return cache[index]

    def differs(self, key):
        left = self._identifier(self.target_styles, self.target.effective_style(key), self._target_ids)
        right = self._identifier(self.source_styles, self.source.effective_style(key), self._source_ids)
        return left != right

    def keys(self):
        # A book uses few xf indices; resolve each once and compare plain
        # identifiers per cell. Only inherited styles need differs().
        target_styles, source_styles = self.target.cell_styles, self.source.cell_styles
        target_cells, source_cells = self.target.cells, self.source.cells
        target_ids = {index: self._identifier(self.target_styles, index, self._target_ids)
                      for index in {0, *target_styles.values()}}
        source_ids = {index: self._identifier(self.source_styles, index, self._source_ids)
                      for index in {0, *source_styles.values()}}
        differs = self.differs
        for key, index in target_styles.items():
            other = source_styles.get(key)
            if other is not None:
                if target_ids[index] != source_ids[other]:
                    yield key
            elif key in source_cells:
                if target_ids[index] != source_ids[0]:
                    yield key
            elif differs(key):
                yield key
        for key, index in source_styles.items():
            if key in target_styles:
                continue
            if key in target_cells:
                if target_ids[0] != source_ids[index]:
                    yield key
            elif differs(key):
                yield key
        # Values on both sides with no entry in cell_styles are explicit 0.
        if target_ids[0] != source_ids[0]:
            for key in target_cells:
                if key in source_cells and key not in target_styles and key not in source_styles:
                    yield key
        for cells, other in ((target_cells, source_cells), (source_cells, target_cells)):
            for key in cells:
                if key not in other and key not in target_styles and key not in source_styles and differs(key):
                    yield key

    def texts(self, key):
        if not self.differs(key):
            return None
        indices = (self.target.effective_style(key), self.source.effective_style(key))
        if indices in self._texts:
            return self._texts[indices]
        texts, canonical = [], []
        for book, styles, names in ((self.target, self.target_styles, self._display_names[0]), (self.source, self.source_styles, self._display_names[1])):
            index = book.effective_style(key)
            texts.append(_description(styles.fields(index), self.reference, names))
            canonical.append(styles.canonical(index))
        if texts[0] == texts[1]:
            texts = [text + " [" + hashlib.sha256(repr(fields).encode("utf-8")).hexdigest()[:8] + "]" for text, fields in zip(texts, canonical)]
        self._texts[indices] = tuple(texts)
        return self._texts[indices]

    def projection(self):
        coordinates = {sheet.id: [] for sheet in self.target.sheets}
        for key in self.keys():
            coordinates[key[0]].append(key[1])
        baseline, changed = {}, {}
        for identifier, values in coordinates.items():
            values.sort(key=coordinate_key)
            left, right = [], []
            for coordinate in values:
                texts = self.texts((identifier, coordinate))
                left.append(_cell_record(coordinate, Cell(FORMAT_KIND, texts[0])) + "\n")
                right.append(_cell_record(coordinate, Cell(FORMAT_KIND, texts[1])) + "\n")
            # Every record differs, so lazygit hunks and WinMerge differences
            # would merge into one block. An unchanged blank line keeps each
            # cell's format a block of its own.
            baseline[identifier], changed[identifier] = "\n".join(left), "\n".join(right)
        return baseline, changed, coordinates

    def unsupported_reason(self, key):
        if self._importer is None:
            self._importer = _StyleImporter(self.target_styles, self.source_styles)
        index = self.source.effective_style(key)
        if index not in self._problems:
            self._problems[index] = self._importer.problem(index)
        return self._problems[index]


class _StyleImporter:
    def __init__(self, target_styles, source_styles):
        self.target, self.source = target_styles, source_styles
        self.modified = False
        self.tree = None
        self._indices, self._exact, self._components, self._styles = {}, None, {}, None
        self._next_numfmt = None

    def _exact_indices(self):
        if self._exact is None:
            self._exact = {}
            for index in range(len(self.target.cellXfs)):
                self._exact.setdefault(self.target.exact(index), index)
        return self._exact

    def problem(self, index):
        source, target = self.source, self.target
        if 0 <= index < len(source.cellXfs) and any(_local(node) == "extLst" for node in source.cellXfs[index]):
            return "書式の拡張機能"
        if target.root is None or any(name not in target.containers for name in ("fonts", "fills", "borders", "cellXfs")):
            return "取り込み先に書式の定義がない"
        # Existing equivalent entries need no copied components or palette.
        if 0 <= index < len(source.cellXfs) and source.exact(index) in self._exact_indices():
            return None
        reason = source.problem(index)
        if reason:
            return reason
        if source.palette != target.palette and source.uses_low_indexed(index):
            return "ブックのパレットが異なる"
        xf = source.cellXfs[index]
        identity = source.style_identity(xf)
        if identity not in target.style_names() and identity[0] != "unnamed":
            node = source._style_elements.get(_index(xf.get("xfId")))
            name = node.get("name", "").casefold() if node is not None else ""
            if any(other.get("name", "").casefold() == name for other in target.cellStyles):
                return "同じ名前で中身の違うスタイル"
        if len(target.cellXfs) >= 64000:
            return "セルの書式の数が 64000 を超える"
        if identity not in target.style_names() and "cellStyleXfs" not in target.containers and (identity[0] != "unnamed" or source.cellStyleXfs):
            return "取り込み先にセルのスタイルの定義がない"
        if len(target.cellStyleXfs) >= 64000 and identity not in target.style_names():
            return "セルのスタイルの数が 64000 を超える"
        # A locale-dependent builtin cannot be overwritten by a custom code.
        for node in [xf] + ([source.cellStyleXfs[_index(xf.get("xfId"))]] if _index(xf.get("xfId")) < len(source.cellStyleXfs) else []):
            fmt = source.numfmt(node.get("numFmtId"))
            if fmt[0] == "builtin" and fmt[1] in target.numFmts:
                return "組み込みの表示形式番号が別の定義に使われている"
        return None

    def _start(self):
        if self.tree is None:
            self.tree = copy.deepcopy(self.target.tree)
            self.root = self.tree.getroot()
            self.containers = {_local(node): node for node in self.root}

    def _container(self, name):
        self._start()
        if name not in self.containers:
            node = etree.Element(_tag(self.target.ns, name), count="0")
            position = _STYLE_CONTAINER_ORDER.index(name)
            following = next((child for child in self.root if _local(child) in _STYLE_CONTAINER_ORDER and _STYLE_CONTAINER_ORDER.index(_local(child)) > position), None)
            self.root.insert(self.root.index(following), node) if following is not None else self.root.append(node)
            self.containers[name] = node
        return self.containers[name]

    def _append(self, name, node):
        container = self._container(name)
        if name in ("cellXfs", "cellStyleXfs") and len(container) >= 64000:
            raise WorkbookError("セルの書式またはスタイルの数が 64000 を超えます")
        index = len(container)
        container.append(node)
        if "count" in container.attrib:
            container.set("count", str(len(container)))
        self.modified = True
        return index

    def _numfmt(self, value):
        fmt = self.source.numfmt(value)
        if fmt[0] == "builtin":
            return fmt[1]
        code = fmt[1]
        container = self.containers.get("numFmts")
        if container is not None:
            found = next((node for node in container if node.get("formatCode") == code), None)
            if found is not None:
                return _index(found.get("numFmtId"))
        for number, text in _INVARIANT_NUMFMTS.items():
            if text == code and (container is None or all(_index(node.get("numFmtId")) != number for node in container)):
                return number
        if self._next_numfmt is None:
            numbers = [163]
            numbers.extend(_index(value, -1) for node in self.target.root.iter() for name, value in node.attrib.items() if _local_attribute(name) == "numFmtId")
            for path, data in self.target.parts.items():
                if path.startswith(("xl/pivotCache/", "xl/pivotTables/")) and path.endswith(".xml"):
                    tree = _xml(data, path)
                    numbers.extend(_index(value, -1) for node in tree.iter() for name, value in node.attrib.items() if _local_attribute(name) == "numFmtId")
            self._next_numfmt = max(numbers) + 1
        number = self._next_numfmt
        self._next_numfmt += 1
        original = self.source.numFmts.get(_index(value))
        node = copy.deepcopy(original) if original is not None else etree.Element(_tag(self.target.ns, "numFmt"), formatCode=code)
        node.set("numFmtId", str(number))
        self._append("numFmts", node)
        return number

    def _component(self, name, number, flatten):
        if name not in self._components:
            mapping = {}
            for index, node in enumerate(getattr(self.target, name)):
                mapping.setdefault(tuple(sorted(flatten(node).items())), index)
            self._components[name] = mapping
        original = getattr(self.source, name)[number]
        key = tuple(sorted(flatten(original).items()))
        mapping = self._components[name]
        if key not in mapping:
            mapping[key] = self._append(name, copy.deepcopy(original))
        return mapping[key]

    def _remap(self, original):
        node = copy.deepcopy(original)
        for name, number in (
            ("numFmtId", self._numfmt(original.get("numFmtId"))),
            ("fontId", self._component("fonts", _index(original.get("fontId")), _font_fields)),
            ("fillId", self._component("fills", _index(original.get("fillId")), _fill_fields)),
            ("borderId", self._component("borders", _index(original.get("borderId")), _border_fields)),
        ):
            if name in node.attrib or number:
                node.set(name, str(number))
        return node

    def _style(self, xf):
        if self._styles is None:
            self._styles = {identity: index for index, identity in self.target._names.items()}
        identity = self.source.style_identity(xf)
        if identity in self._styles:
            return self._styles[identity]
        source_index = _index(xf.get("xfId"))
        if identity[0] == "unnamed" and not self.source.cellStyleXfs and not self.target.cellStyleXfs:
            return 0
        if source_index < len(self.source.cellStyleXfs):
            original = self.source.cellStyleXfs[source_index]
        else:
            original = etree.Element(_tag(self.target.ns, "xf"))
        node = self._remap(original)
        number = self._append("cellStyleXfs", node)
        if identity[0] != "unnamed":
            original = self.source._style_elements[source_index]
            style = etree.Element(_tag(self.target.ns, "cellStyle"))
            for name in ("name", "builtinId", "customBuiltin", "iLevel", "hidden"):
                if original.get(name) is not None:
                    style.set(name, original.get(name))
            style.set("xfId", str(number))
            self._append("cellStyles", style)
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
            self._start()
            original = self.source.cellXfs[index]
            node = self._remap(original)
            number = self._style(original)
            if "xfId" in node.attrib or number:
                node.set("xfId", str(number))
            result = self._append("cellXfs", node)
            self._exact_indices().setdefault(self.source.exact(index), result)
        self._indices[index] = result
        return result

    def serialize(self, original_bytes):
        return _serialize(self.tree, original_bytes) if self.modified else original_bytes


def _local_attribute(name):
    return name.rpartition("}")[2]


def _check_styles_appended(original_bytes, result_bytes):
    failure = "作り直したブックの書式の検証に失敗しました。書き込みは行っていません"
    before, after = _xml(original_bytes, "styles.xml").getroot(), _xml(result_bytes, "styles.xml").getroot()
    if before.tag != after.tag or before.attrib != after.attrib or before.nsmap != after.nsmap:
        raise WorkbookError(failure)
    old = {_local(node): node for node in before}
    new = {_local(node): node for node in after}
    if len(old) != len(before) or len(new) != len(after) or not set(old) <= set(new) or set(new) - set(old) - {"numFmts", "cellStyles"}:
        raise WorkbookError(failure)
    grown = {"numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles"}
    if [_local(node) for node in after if _local(node) in old] != [_local(node) for node in before]:
        raise WorkbookError(failure)
    for name in set(new) - set(old):
        ordered = [_STYLE_CONTAINER_ORDER.index(_local(node)) for node in after if _local(node) in _STYLE_CONTAINER_ORDER]
        if ordered != sorted(ordered) or new[name].get("count") != str(len(new[name])):
            raise WorkbookError(failure)
    for name, node in old.items():
        result = new[name]
        if name in grown and len(result) > len(node) and "count" in result.attrib and result.get("count") != str(len(result)):
            raise WorkbookError(failure)
        if name not in grown:
            if etree.tostring(node) != etree.tostring(result):
                raise WorkbookError(failure)
        elif (
            node.tag != result.tag
            or {k: v for k, v in node.attrib.items() if k != "count"} != {k: v for k, v in result.attrib.items() if k != "count"}
            or len(node) > len(result)
            or any(etree.tostring(a) != etree.tostring(b) for a, b in zip(node, result))
        ):
            raise WorkbookError(failure)


def _column_styles(nodes):
    """Sorted (min, max, style) of valid, nonzero <col> styles."""
    entries = []
    for node in nodes:
        first, last, style = (_index(node.get(name), -1) for name in ("min", "max", "style"))
        if 1 <= first <= last <= 16384 and style > 0:
            entries.append((first, last, style))
    entries.sort()
    return entries


def _column_style(entries, column):
    position = bisect_right(entries, (column, 16384, float("inf"))) - 1
    if position >= 0 and column <= entries[position][1]:
        return entries[position][2]
    return 0


def _tree_inherited_style(data, row, column):
    if row is not None and _bool(row.get("customFormat")) is True:
        return _index(row.get("s"))
    # A worksheet may hold several <cols> elements; resolve them as the loader does.
    root = data.getparent()
    return _column_style(_column_styles(col for cols in _children(root, "cols") for col in _children(cols, "col")), column)


class Workbook:
    def __init__(self, data):
        self.data = data
        self._reset()
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as archive:
                self._infos = archive.infolist()
                if len({info.filename for info in self._infos}) != len(self._infos):
                    raise WorkbookError("同名パーツが重複する Excel は扱えません")
                self._parts = {info.filename: archive.read(info) for info in self._infos}
                self._comment = archive.comment
            if any(path.startswith("_xmlsignatures/") for path in self._parts):
                raise WorkbookError("電子署名付きの Excel はセル単位で反映できません")
            self._load()
        except WorkbookError:
            raise
        except (OSError, KeyError, RuntimeError, zipfile.BadZipFile, ValueError) as error:
            raise WorkbookError("Excel ブックを読めません: " + str(error)) from error

    def _reset(self):
        self.cells, self.cell_styles, self._trees, self._unsafe = {}, {}, {}, {}
        self._ranges, self._merges, self._order, self._positions = {}, {}, {}, {}
        self._row_styles, self._col_styles, self._c_counts = {}, {}, {}
        self._stylesheet = None

    def _read_xml(self, path):
        if path not in self._parts:
            raise WorkbookError("Excel のパーツがありません: " + path)
        tree = _xml(self._parts[path], path)
        self._trees[path] = tree
        return tree

    def _load(self, load_sheets=True):
        root_relationships = self._read_xml("_rels/.rels").getroot()
        office = [node for node in root_relationships if node.get("Type", "").endswith("/officeDocument")]
        if len(office) != 1 or office[0].get("TargetMode") == "External":
            raise WorkbookError("Excel の workbook 参照が不正です")
        self._workbook_path = _target("", office[0].get("Target", ""))
        self._workbook_rels = _rels_path(self._workbook_path)
        workbook = self._read_xml(self._workbook_path).getroot()
        if _local(workbook) != "workbook":
            raise WorkbookError("Excel の workbook XML が不正です")
        self._ns = _namespace(workbook)
        relationships = self._read_xml(self._workbook_rels).getroot()
        rels = {}
        for node in relationships:
            identifier = node.get("Id")
            if identifier in rels:
                raise WorkbookError("Excel の relationship ID が重複しています")
            rels[identifier] = node
        style_rels = [node for node in relationships if node.get("Type", "").endswith("/styles")]
        self._styles_ambiguous = len(style_rels) > 1
        self._styles_path = _target(self._workbook_path, style_rels[0].get("Target", "")) if len(style_rels) == 1 and style_rels[0].get("TargetMode") != "External" else None
        strings = []
        self._rich_strings = set()
        string_rels = [node for node in relationships if node.get("Type", "").endswith("/sharedStrings")]
        if len(string_rels) > 1:
            raise WorkbookError("Excel の sharedStrings が重複しています")
        if string_rels:
            path = _target(self._workbook_path, string_rels[0].get("Target", ""))
            for node in self._read_xml(path).getroot():
                if _local(node) == "si":
                    if any(_local(child) not in _PLAIN_STRING for child in node):
                        self._rich_strings.add(len(strings))
                    strings.append(_string(node))
        self._strings = strings
        sheets = _child(workbook, "sheets")
        if sheets is None:
            raise WorkbookError("Excel に sheets 要素がありません")
        descriptors = []
        structure = []
        names, identifiers = set(), set()
        for node in sheets:
            if _local(node) != "sheet":
                continue
            identifier, name = node.get("sheetId"), node.get("name")
            if not identifier or not name or identifier in identifiers or name in names:
                raise WorkbookError("シート名または sheetId が不正です")
            names.add(name)
            identifiers.add(identifier)
            relationship = next((value for key, value in node.attrib.items() if etree.QName(key).localname == "id"), None)
            rel = rels.get(relationship)
            if rel is None or rel.get("TargetMode") == "External":
                raise WorkbookError("シートの参照が不正です: " + name)
            path = _target(self._workbook_path, rel.get("Target", ""))
            kind = rel.get("Type", "").rsplit("/", 1)[-1]
            structure.append((identifier, name, relationship, path, kind))
            if kind != "worksheet":
                continue  # Chartsheets and their original parts stay untouched.
            if path not in self._parts:
                raise WorkbookError("Excel のパーツがありません: " + path)
            sheet = Sheet(identifier, name, path)
            descriptors.append(sheet)
            if load_sheets:
                self._load_sheet(sheet, strings)
        self.sheets = tuple(descriptors)
        properties = _child(workbook, "workbookPr")
        date_value = properties.get("date1904", "0") if properties is not None else "0"
        if date_value not in ("0", "1", "false", "true"):
            raise WorkbookError("Excel の日付基準が不正です")
        date1904 = date_value in ("1", "true")
        names_node = _child(workbook, "definedNames")
        defined_names = () if names_node is None else tuple(
            (tuple(sorted(node.attrib.items())), node.text or "") for node in names_node if _local(node) == "definedName"
        )
        external = []
        for node in relationships:
            if not node.get("Type", "").endswith("/externalLink"):
                continue
            path = _target(self._workbook_path, node.get("Target", ""))
            link = self._read_xml(path).getroot()
            definitions = tuple((_local(value), tuple(sorted(value.attrib.items()))) for value in link)
            rel_path = _rels_path(path)
            targets = ()
            if rel_path in self._parts:
                link_relationships = self._read_xml(rel_path).getroot()
                targets = tuple((value.get("Id"), value.get("Type"), value.get("Target"), value.get("TargetMode")) for value in link_relationships)
            external.append((node.get("Id"), node.get("Type"), path, node.get("TargetMode"), definitions, targets))
        external = tuple(external)
        self._structure = (self._workbook_path, self._ns, tuple(structure), date1904, defined_names, external)

    def _load_sheet(self, sheet, strings, expected=None, expected_format=None):
        # Only a row's XML is retained during loading. The original part bytes
        # remain available for rebuilding a selected sheet later.
        namespace = self._ns
        row_tag, cell_tag, value_tag, formula_tag, inline_tag = (
            _tag(namespace, name) for name in ("row", "c", "v", "f", "is")
        )
        data_tag, merges_tag, text_tag = (_tag(namespace, name) for name in ("sheetData", "mergeCells", "t"))
        cols_tag = _tag(namespace, "cols")
        self._row_styles[sheet.id], self._col_styles[sheet.id] = {}, []
        row_styles, cell_styles, c_count = self._row_styles[sheet.id], self.cell_styles, 0
        style_indices = {"0": 0}
        self._ranges[sheet.id] = []
        self._merges[sheet.id] = []
        coordinates = set()
        order, positions = [], []
        monotonic, previous, count = True, -1, 0
        root, data_seen, merges_seen = None, False, False
        try:
            context = etree.iterparse(
                io.BytesIO(self._parts[sheet.path]), events=("end",),
                tag=(row_tag, "{*}row", "row", data_tag, merges_tag, cols_tag),
                resolve_entities=False, no_network=True, remove_blank_text=False,
            )
            for _, element in context:
                if root is None:
                    root = element.getroottree().getroot()
                parent = element.getparent()
                tag = element.tag
                if tag == data_tag and parent is root:
                    if data_seen:
                        raise WorkbookError("XML の要素が重複しています: sheetData")
                    data_seen = True
                elif tag == cols_tag and parent is root:
                    self._col_styles[sheet.id] = sorted(self._col_styles[sheet.id] + _column_styles(_children(element, "col")))
                elif tag == merges_tag and parent is root:
                    if merges_seen:
                        raise WorkbookError("XML の要素が重複しています: mergeCells")
                    merges_seen = True
                    self._merges[sheet.id] = [
                        _range(node.get("ref", "")) for node in element if _local(node) == "mergeCell"
                    ]
                elif (tag == row_tag or _local(element) == "row") and parent is not None and parent.tag == data_tag and parent.getparent() is root:
                    if element.get("customFormat") is not None and _bool(element.get("customFormat")) is True:
                        row_number = _index(element.get("r"), -1)
                        if row_number < 1:
                            first_cell = next((node for node in element if _local(node) == "c"), None)
                            if first_cell is not None:
                                row_number = coordinate_key(first_cell.get("r", ""))[0]
                        if row_number > 0:
                            row_styles[row_number] = _index(element.get("s"))
                    for node in element:
                        if node.tag != cell_tag and _local(node) != "c":
                            continue
                        coordinate = node.get("r", "")
                        row, column = coordinate_key(coordinate)
                        if coordinate in coordinates:
                            raise WorkbookError("セル番地が重複しています: " + sheet.name + "!" + coordinate)
                        coordinates.add(coordinate)
                        key = (sheet.id, coordinate)
                        c_count += 1
                        style_text = node.get("s", "0")
                        if style_text not in style_indices:
                            style_indices[style_text] = _index(style_text)
                        style = style_indices[style_text]
                        unsafe = None
                        if any(name not in ("r", "s", "t") for name in node.attrib):
                            unsafe = "セルのメタデータ"
                        # Unusual cell namespaces retain the original local-name
                        # handling; ordinary cells use precomputed Clark tags.
                        if node.tag == cell_tag:
                            v_tag, f_tag, is_tag = value_tag, formula_tag, inline_tag
                        else:
                            cell_namespace = _namespace(node)
                            v_tag, f_tag, is_tag = (_tag(cell_namespace, name) for name in ("v", "f", "is"))
                        value_node = formula = inline = None
                        inline_count = 0
                        for child in node:
                            child_tag = child.tag
                            if child_tag == v_tag:
                                if value_node is not None:
                                    raise WorkbookError("XML の要素が重複しています: v")
                                value_node = child
                            elif child_tag == f_tag:
                                if formula is not None:
                                    raise WorkbookError("XML の要素が重複しています: f")
                                formula = child
                            elif child_tag == is_tag:
                                inline = child
                                inline_count += 1
                            if isinstance(child_tag, str) and child_tag not in (value_tag, formula_tag, inline_tag):
                                unsafe = "セルの拡張機能"
                        value = value_node.text if value_node is not None else None
                        kind = node.get("t", "n")
                        cell = None
                        if formula is not None:
                            formula_type = formula.get("t", "normal")
                            if formula_type not in ("", "normal"):
                                unsafe = "共有・配列・データテーブル数式"
                                if formula.get("ref"):
                                    self._ranges[sheet.id].append(_range(formula.get("ref")))
                            if any(name not in ("t", "ca") for name in formula.attrib):
                                unsafe = "数式の拡張機能"
                            cell = Cell("formula", "=" + (formula.text or ""))
                        elif kind == "inlineStr":
                            if inline_count > 1:
                                raise WorkbookError("XML の要素が重複しています: is")
                            if inline is not None:
                                if len(inline) == 1 and inline[0].tag == text_tag and len(inline[0]) == 0:
                                    cell = Cell("string", inline[0].text or "")
                                else:
                                    if any(_local(child) not in _PLAIN_STRING for child in inline):
                                        unsafe = "リッチテキスト"
                                    cell = Cell("string", _string(inline))
                        elif kind == "s" and value is not None:
                            try:
                                number = int(value)
                                if number < 0:
                                    raise IndexError
                                cell = Cell("string", strings[number])
                                if number in self._rich_strings:
                                    unsafe = "リッチテキスト"
                            except (ValueError, IndexError):
                                raise WorkbookError("sharedStrings の参照が不正です: " + coordinate) from None
                        elif kind in ("str", "e", "d") and value_node is not None:
                            cell = Cell({"str": "string", "e": "error", "d": "date"}[kind], value or "")
                        elif kind == "b" and value is not None:
                            if value not in ("0", "1", "true", "false"):
                                raise WorkbookError("真偽値セルが不正です: " + coordinate)
                            cell = Cell("boolean", value in ("1", "true"))
                        elif kind == "n" and value is not None:
                            cell = Cell("number", _number(value))
                        elif kind not in ("inlineStr", "s", "str", "e", "d", "b", "n"):
                            unsafe = "未知のセル型 " + kind
                            cell = Cell("string", value or "")
                        if cell is not None:
                            count += 1
                            if expected is None:
                                self.cells[key] = cell
                                position = row * 16384 + column - 1
                                monotonic = monotonic and position > previous
                                previous = position
                                order.append(key)
                                positions.append(position)
                            elif cell != expected(key):
                                raise WorkbookError("作り直したブックの検証に失敗しました。書き込みは行っていません")
                        if expected_format is not None:
                            expected_format(key, style)
                        elif cell is None or style != 0:
                            cell_styles[key] = style
                        if unsafe:
                            self._unsafe[key] = unsafe
                    element.clear()
                    while element.getprevious() is not None:
                        parent.remove(element.getprevious())
            root = context.root
            if root.getroottree().docinfo.doctype:
                raise WorkbookError("DTD を含む XML は扱えません: " + sheet.path)
        except (etree.XMLSyntaxError, ValueError) as error:
            if isinstance(error, WorkbookError):
                raise
            raise WorkbookError("Excel の XML を読めません: " + sheet.path) from error
        self._c_counts[sheet.id] = c_count
        if _local(root) != "worksheet" or _namespace(root) != namespace:
            raise WorkbookError("worksheet XML が不正です: " + sheet.name)
        if not data_seen:
            raise WorkbookError("sheetData がありません: " + sheet.name)
        if expected is None:
            if not monotonic:
                indices = sorted(range(len(order)), key=positions.__getitem__)
                order = [order[index] for index in indices]
                positions = [positions[index] for index in indices]
            self._order[sheet.id] = order
            self._positions[sheet.id] = array("Q", positions)
        return count

    def stylesheet(self):
        if self._stylesheet is None:
            if self._styles_ambiguous:
                raise WorkbookError("Excel の styles が重複しています")
            self._stylesheet = _Styles(self._parts.get(self._styles_path), self._styles_path)
            self._stylesheet.parts = self._parts
        return self._stylesheet

    def effective_style(self, key):
        if key in self.cell_styles:
            return self.cell_styles[key]
        if key in self.cells:
            return 0
        row, column = coordinate_key(key[1])
        return self._inherited_style(key[0], row, column)

    def _inherited_style(self, identifier, row, column):
        rows = self._row_styles.get(identifier, {})
        if row in rows:
            return rows[row]
        return _column_style(self._col_styles.get(identifier, []), column)

    def compare_formats(self, other):
        return FormatComparison(self, other)

    def assert_compatible(self, other):
        if self._structure != other._structure:
            labels = ("ブックの内部構成", "ブックの内部構成", "シートの名前・順序・内部の参照", "日付基準", "名前定義", "外部参照")
            differences = list(dict.fromkeys(label for label, left, right in
                                             zip(labels, self._structure, other._structure) if left != right))
            raise WorkbookError("・".join(differences) + "が違うブックはセル単位で反映できません")

    def _sync_order(self):
        # cells remains a public dict. Rebuild the cache only if a caller has
        # added, removed or replaced keys after loading the workbook.
        if len(self.cells) == sum(map(len, self._order.values())) and all(
            key in self.cells for keys in self._order.values() for key in keys
        ):
            return
        orders = {sheet.id: [] for sheet in self.sheets}
        for key in self.cells:
            row, column = coordinate_key(key[1])
            orders[key[0]].append((row * 16384 + column - 1, key))
        for identifier, values in orders.items():
            values.sort()
            self._order[identifier] = [key for _, key in values]
            self._positions[identifier] = array("Q", (position for position, _ in values))

    def projection(self, other=None):
        if other is not None:
            self.assert_compatible(other)
            other._sync_order()
        self._sync_order()
        lines = {}
        for sheet in self.sheets:
            keys = self._order[sheet.id]
            if other is not None and keys != other._order[sheet.id]:
                # Coordinates are parsed and sorted at load time, not once per
                # projection. Merge the two numeric orders without regex work.
                left, right = keys, other._order[sheet.id]
                left_positions, right_positions = self._positions[sheet.id], other._positions[sheet.id]
                keys, i, j = [], 0, 0
                while i < len(left) and j < len(right):
                    a, b = left_positions[i], right_positions[j]
                    if a <= b:
                        keys.append(left[i])
                        i += 1
                    if b <= a:
                        if b < a:
                            keys.append(right[j])
                        j += 1
                keys.extend(left[i:])
                keys.extend(right[j:])
            output = io.StringIO()
            for key in keys:
                output.write(_cell_record(key[1], self.cells.get(key)))
                output.write("\n")
            lines[sheet.id] = output.getvalue()
        return lines

    def unsupported_reason(self, key):
        """Return the same reason used by rebuild, or None for an ordinary cell."""
        identifier, coordinate = key
        if identifier not in {sheet.id for sheet in self.sheets}:
            raise WorkbookError("セル選択に未知のシートがあります")
        coordinate_key(coordinate)
        reason = self._unsafe.get(key)
        if reason:
            return reason
        if any(_inside(coordinate, rectangle) for rectangle in self._ranges[identifier]):
            return "配列・共有・データテーブル数式の範囲"
        for rectangle in self._merges[identifier]:
            if _inside(coordinate, rectangle) and coordinate_key(coordinate) != rectangle[0]:
                return "結合セルの左上以外"
        return None

    def _check_selected(self, selected):
        sheets = {sheet.id: sheet for sheet in self.sheets}
        for key in selected:
            reason = self.unsupported_reason(key)
            if reason:
                raise WorkbookError(sheets[key[0]].name + "!" + key[1] + " はセル単位で反映できません: " + reason)

    def rebuild(self, source, selected, formats=()):
        self.assert_compatible(source)
        selected, formats = set(selected), set(formats)
        self._check_selected(selected)
        source._check_selected(selected)
        changed = {key for key in selected if self.cells.get(key) != source.cells.get(key)}
        comparison = self.compare_formats(source) if formats else None
        sheets = {sheet.id: sheet for sheet in self.sheets}
        for key in formats:
            if key[0] not in sheets:
                raise WorkbookError("セル選択に未知のシートがあります")
            coordinate_key(key[1])
            reason = comparison.unsupported_reason(key)
            if reason:
                raise WorkbookError(sheets[key[0]].name + "!" + key[1] + " の書式は反映できません: " + reason)
        restyled = {key for key in formats if comparison.differs(key)}
        if not changed and not restyled:
            return self.data
        updates, removed, styles = {}, set(), {}
        if restyled:
            importer = _StyleImporter(comparison.target_styles, comparison.source_styles)
            for key in restyled:
                styles[key] = importer.index_for(source.effective_style(key))
            if importer.modified:
                updates[self._styles_path] = importer.serialize(self._parts[self._styles_path])
        for sheet in self.sheets:
            coordinates = [coordinate for identifier, coordinate in changed if identifier == sheet.id]
            formats_for_sheet = {coordinate: index for (identifier, coordinate), index in styles.items() if identifier == sheet.id}
            if coordinates or formats_for_sheet:
                updates[sheet.path] = self._rebuild_sheet(sheet, source, coordinates, formats_for_sheet)
        if changed:
            self._recalculate(updates, removed)
        output = io.BytesIO()
        with zipfile.ZipFile(output, "w") as archive:
            archive.comment = self._comment
            for info in self._infos:
                if info.filename not in removed:
                    archive.writestr(copy.copy(info), updates.get(info.filename, self._parts[info.filename]))
        result = output.getvalue()
        self._validate_result(result, source, changed, updates, removed, restyled)
        return result

    def _rebuild_sheet(self, sheet, source, coordinates, styles=None):
        # Parse the original bytes only for a sheet being edited. Keeping the
        # tree in this helper releases it before the result is verified.
        tree = _xml(self._parts[sheet.path], sheet.path)
        root = tree.getroot()
        data = _child(root, "sheetData")
        rows = {node.get("r"): node for node in _children(data, "row")}
        styles = styles or {}
        values = set(coordinates)
        touched = values | styles.keys()
        for coordinate in sorted(touched, key=coordinate_key):
            if coordinate in values:
                self._replace(data, rows, coordinate, source.cells.get((sheet.id, coordinate)))
            if coordinate in styles:
                self._set_style(data, rows, coordinate, styles[coordinate])
        dimension = _child(root, "dimension")
        if dimension is not None:
            first, last = _range(dimension.get("ref", "A1"))
            present = [coordinate_key(coordinate) for coordinate in touched if self._cell_element(data, rows, coordinate, False) is not None]
            if present:
                first = (min(first[0], *(row for row, _ in present)), min(first[1], *(column for _, column in present)))
                last = (max(last[0], *(row for row, _ in present)), max(last[1], *(column for _, column in present)))
                dimension.set("ref", self._coordinate(first) + ":" + self._coordinate(last) if first != last else self._coordinate(first))
        return _serialize(tree, self._parts[sheet.path])

    def _validate_result(self, result, source, changed, updates, removed, restyled=None):
        failure = "作り直したブックの検証に失敗しました。書き込みは行っていません"
        restyled = restyled or set()
        parts = {path: content for path, content in self._parts.items() if path not in removed}
        parts.update(updates)
        # Check the actual ZIP, including every untouched/removed part. Once its
        # bytes match, unchanged worksheets need no second cell-level parse.
        try:
            with zipfile.ZipFile(io.BytesIO(result)) as archive:
                names = archive.namelist()
                if len(names) != len(set(names)) or set(names) != set(parts) or archive.comment != self._comment:
                    raise WorkbookError(failure)
                if any(name.startswith("_xmlsignatures/") for name in names):
                    raise WorkbookError(failure)
                for name in names:
                    if archive.read(name) != parts[name]:
                        raise WorkbookError(failure)
        except (OSError, KeyError, RuntimeError, zipfile.BadZipFile, ValueError) as error:
            if isinstance(error, WorkbookError):
                raise
            raise WorkbookError(failure) from error
        if self._styles_path in updates:
            _check_styles_appended(self._parts[self._styles_path], updates[self._styles_path])
        actual = Workbook.__new__(Workbook)
        actual._reset()
        actual._parts = parts
        # Reload workbook/relationship metadata and retain the same structural
        # compatibility check as a full Workbook(result) validation.
        actual._load(load_sheets=False)
        self.assert_compatible(actual)
        if actual._strings != self._strings or actual._rich_strings != self._rich_strings:
            raise WorkbookError(failure)
        actual_styles = actual.stylesheet() if restyled else None
        source_styles = source.stylesheet() if restyled else None
        cell_styles, cells = self.cell_styles, self.cells

        def expected(key):
            return source.cells.get(key) if key in changed else self.cells.get(key)

        for sheet in self.sheets:
            if sheet.path not in updates:
                continue
            touched = {key for key in changed | restyled if key[0] == sheet.id}
            seen = set()

            def check(key, index):
                if key in restyled:
                    if actual_styles.exact(index) != source_styles.exact(source.effective_style(key)):
                        raise WorkbookError(failure)
                    seen.add(key)
                    return
                expected_style = cell_styles.get(key)
                if expected_style is None:
                    if key in cells:
                        expected_style = 0
                    elif key in touched:
                        expected_style = self.effective_style(key)
                    else:
                        # A <c> appeared outside the selection.
                        raise WorkbookError(failure)
                if index != expected_style:
                    raise WorkbookError(failure)
                if key in touched:
                    seen.add(key)

            expected_count = sum(1 for identifier, _ in self.cells if identifier == sheet.id)
            for key in changed:
                if key[0] == sheet.id:
                    expected_count += (key in source.cells) - (key in self.cells)
            count = actual._load_sheet(sheet, actual._strings, expected=expected, expected_format=check)
            if count != expected_count or actual._row_styles[sheet.id] != self._row_styles[sheet.id] or actual._col_styles[sheet.id] != self._col_styles[sheet.id]:
                raise WorkbookError(failure)
            present_count = len(seen)
            for key in touched - seen:
                row, column = coordinate_key(key[1])
                check(key, actual._inherited_style(key[0], row, column))
            old_touched = sum(key in self.cells or key in self.cell_styles for key in touched)
            if actual._c_counts[sheet.id] != self._c_counts[sheet.id] - old_touched + present_count:
                raise WorkbookError(failure)

    @staticmethod
    def _coordinate(position):
        row, column = position
        letters = ""
        while column:
            column, remainder = divmod(column - 1, 26)
            letters = chr(65 + remainder) + letters
        return letters + str(row)

    def _cell_element(self, data, rows, coordinate, create):
        row_number, column_number = coordinate_key(coordinate)
        row = rows.get(str(row_number))
        if row is None:
            if not create:
                return None
            row = etree.Element(_tag(self._ns, "row"), r=str(row_number))
            following = next((node for node in _children(data, "row") if int(node.get("r", "0")) > row_number), None)
            rows[str(row_number)] = row
            if following is None:
                data.append(row)
            else:
                data.insert(data.index(following), row)
        nodes = _children(row, "c")
        node = next((node for node in nodes if node.get("r") == coordinate), None)
        if node is None and create:
            node = etree.Element(_tag(self._ns, "c"), r=coordinate)
            following = next((value for value in nodes if coordinate_key(value.get("r"))[1] > column_number), None)
            if following is None:
                extension = next((value for value in row if isinstance(value.tag, str) and _local(value) != "c"), None)
                if extension is None:
                    row.append(node)
                else:
                    row.insert(row.index(extension), node)
            else:
                row.insert(row.index(following), node)
        return node

    def _set_style(self, data, rows, coordinate, index):
        row_number, column_number = coordinate_key(coordinate)
        inherited = _tree_inherited_style(data, rows.get(str(row_number)), column_number)
        node = self._cell_element(data, rows, coordinate, False)
        if node is None and index == inherited:
            return
        if node is None:
            node = self._cell_element(data, rows, coordinate, True)
        if index:
            node.set("s", str(index))
        else:
            node.attrib.pop("s", None)
        if set(node.attrib) == {"r"} and not len(node) and inherited == 0:
            node.getparent().remove(node)

    def _replace(self, data, rows, coordinate, cell):
        row_number, column_number = coordinate_key(coordinate)
        inherited = _tree_inherited_style(data, rows.get(str(row_number)), column_number)
        node = self._cell_element(data, rows, coordinate, False)
        if node is None:
            if cell is None:
                return
            node = self._cell_element(data, rows, coordinate, True)
            if inherited > 0:
                node.set("s", str(inherited))
        for child in list(node):
            if _local(child) in ("f", "v", "is"):
                node.remove(child)
        node.attrib.pop("t", None)
        if cell is None:
            if set(node.attrib) == {"r"} and not len(node) and inherited == 0:
                node.getparent().remove(node)
            return
        if cell.kind == "formula":
            etree.SubElement(node, _tag(self._ns, "f")).text = cell.value[1:]
        elif cell.kind == "string":
            node.set("t", "inlineStr")
            inline = etree.SubElement(node, _tag(self._ns, "is"))
            text = etree.SubElement(inline, _tag(self._ns, "t"))
            text.set(_XML_SPACE, "preserve")
            text.text = cell.value
        else:
            types = {"number": "n", "boolean": "b", "error": "e", "date": "d"}
            node.set("t", types[cell.kind])
            value = etree.SubElement(node, _tag(self._ns, "v"))
            value.text = ("1" if cell.value else "0") if cell.kind == "boolean" else cell.value

    def _recalculate(self, updates, removed):
        tree = copy.deepcopy(self._trees[self._workbook_path])
        workbook = tree.getroot()
        calculation = _child(workbook, "calcPr")
        if calculation is None:
            calculation = etree.Element(_tag(self._ns, "calcPr"))
            following_names = {"oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"}
            following = next((node for node in workbook if _local(node) in following_names), None)
            if following is None:
                workbook.append(calculation)
            else:
                workbook.insert(workbook.index(following), calculation)
        # Recalculate stale caches once on open. forceFullCalc would persist as
        # Excel's "always fully recalculate" setting, so it is left untouched.
        calculation.set("fullCalcOnLoad", "1")
        updates[self._workbook_path] = _serialize(tree, self._parts[self._workbook_path])
        relationships = copy.deepcopy(self._trees[self._workbook_rels])
        rels = relationships.getroot()
        chains = [node for node in rels if node.get("Type", "").endswith("/calcChain")]
        for node in chains:
            path = _target(self._workbook_path, node.get("Target", ""))
            removed.add(path)
            removed.add(_rels_path(path))
            rels.remove(node)
        if chains:
            updates[self._workbook_rels] = _serialize(relationships, self._parts[self._workbook_rels])
            content = self._read_xml("[Content_Types].xml")
            content = copy.deepcopy(content)
            for node in list(content.getroot()):
                if node.get("PartName", "").lstrip("/") in removed:
                    content.getroot().remove(node)
            updates["[Content_Types].xml"] = _serialize(content, self._parts["[Content_Types].xml"])
