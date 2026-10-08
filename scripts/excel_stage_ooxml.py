"""Lossless package editing for staging ordinary OOXML cell values/formulas.

Only the selected cells are reconstructed.  In particular, this module does
not save a workbook through openpyxl, which can discard unsupported features.
"""

import copy
import io
import json
import posixpath
import re
import zipfile
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation

from lxml import etree


class WorkbookError(ValueError):
    """A workbook or requested edit cannot safely be partially staged."""


@dataclass(frozen=True)
class Cell:
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


def _number(value):
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


def cell_record(coordinate, cell):
    coordinate_key(coordinate)
    values = [coordinate, cell.kind, cell.value] if cell is not None else [coordinate, "blank", None]
    return json.dumps(values, ensure_ascii=False, allow_nan=False, separators=(",", ":"))


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
    if kind not in _KINDS:
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
    return [child for child in node if _local(child) == name and _namespace(child) == _namespace(node)]


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


class Workbook:
    def __init__(self, data):
        self.data = data
        self.cells = {}
        self._trees = {}
        self._unsafe = {}
        self._ranges = {}
        self._merges = {}
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as archive:
                self._infos = archive.infolist()
                if len({info.filename for info in self._infos}) != len(self._infos):
                    raise WorkbookError("同名パーツが重複する Excel は扱えません")
                self._parts = {info.filename: archive.read(info) for info in self._infos}
                self._comment = archive.comment
            if any(path.startswith("_xmlsignatures/") for path in self._parts):
                raise WorkbookError("電子署名付きの Excel は部分ステージできません")
            self._load()
        except WorkbookError:
            raise
        except (OSError, KeyError, RuntimeError, zipfile.BadZipFile, ValueError) as error:
            raise WorkbookError("Excel ブックを読めません: " + str(error)) from error

    def _read_xml(self, path):
        if path not in self._parts:
            raise WorkbookError("Excel のパーツがありません: " + path)
        tree = _xml(self._parts[path], path)
        self._trees[path] = tree
        return tree

    def _load(self):
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
            sheet = Sheet(identifier, name, path)
            descriptors.append(sheet)
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

    def _load_sheet(self, sheet, strings):
        root = self._read_xml(sheet.path).getroot()
        if _local(root) != "worksheet" or _namespace(root) != self._ns:
            raise WorkbookError("worksheet XML が不正です: " + sheet.name)
        data = _child(root, "sheetData")
        if data is None:
            raise WorkbookError("sheetData がありません: " + sheet.name)
        self._ranges[sheet.id] = []
        self._merges[sheet.id] = []
        merges = _child(root, "mergeCells")
        if merges is not None:
            self._merges[sheet.id] = [_range(node.get("ref", "")) for node in merges if _local(node) == "mergeCell"]
        coordinates = set()
        for row in data:
            if _local(row) != "row":
                continue
            for node in row:
                if _local(node) != "c":
                    continue
                coordinate = node.get("r", "")
                coordinate_key(coordinate)
                if coordinate in coordinates:
                    raise WorkbookError("セル番地が重複しています: " + sheet.name + "!" + coordinate)
                coordinates.add(coordinate)
                key = (sheet.id, coordinate)
                unsafe = None
                if any(name not in ("r", "s", "t") for name in node.attrib):
                    unsafe = "セルのメタデータ"
                if any(_local(child) not in ("f", "v", "is") or _namespace(child) != self._ns for child in node if isinstance(child.tag, str)):
                    unsafe = "セルの拡張機能"
                value_node = _child(node, "v")
                formula = _child(node, "f")
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
                    inline = _child(node, "is")
                    if inline is not None:
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
                    self.cells[key] = cell
                if unsafe:
                    self._unsafe[key] = unsafe

    def assert_compatible(self, other):
        if self._structure != other._structure:
            raise WorkbookError("シート構成・名前・日付基準・名前定義・外部参照が変わったブックは部分ステージできません")

    def projection(self, other=None):
        if other is not None:
            self.assert_compatible(other)
        keys = set(self.cells)
        if other is not None:
            keys.update(other.cells)
        lines = {sheet.id: [] for sheet in self.sheets}
        for identifier, coordinate in sorted(keys, key=lambda key: (key[0], coordinate_key(key[1]))):
            lines[identifier].append(cell_record(coordinate, self.cells.get((identifier, coordinate))) + "\n")
        return {identifier: "".join(records) for identifier, records in lines.items()}

    def _check_selected(self, selected):
        sheets = {sheet.id: sheet for sheet in self.sheets}
        for identifier, coordinate in selected:
            if identifier not in sheets:
                raise WorkbookError("セル選択に未知のシートがあります")
            coordinate_key(coordinate)
            reason = self._unsafe.get((identifier, coordinate))
            if reason or any(_inside(coordinate, rectangle) for rectangle in self._ranges[identifier]):
                raise WorkbookError(sheets[identifier].name + "!" + coordinate + " は部分ステージできません: " + (reason or "配列・共有・データテーブル数式の範囲"))
            for rectangle in self._merges[identifier]:
                if _inside(coordinate, rectangle) and coordinate_key(coordinate) != rectangle[0]:
                    raise WorkbookError("結合セルの左上以外は部分ステージできません: " + sheets[identifier].name + "!" + coordinate)

    def rebuild(self, source, selected):
        self.assert_compatible(source)
        selected = set(selected)
        self._check_selected(selected)
        source._check_selected(selected)
        changed = {key for key in selected if self.cells.get(key) != source.cells.get(key)}
        if not changed:
            return self.data
        updates, removed = {}, set()
        for sheet in self.sheets:
            coordinates = [coordinate for identifier, coordinate in changed if identifier == sheet.id]
            if not coordinates:
                continue
            tree = copy.deepcopy(self._trees[sheet.path])
            root = tree.getroot()
            data = _child(root, "sheetData")
            # Index rows once; scanning sheetData per selected cell was
            # quadratic for large sheets with many selected cells.
            rows = {node.get("r"): node for node in _children(data, "row")}
            for coordinate in sorted(coordinates, key=coordinate_key):
                self._replace(data, rows, coordinate, source.cells.get((sheet.id, coordinate)))
            dimension = _child(root, "dimension")
            if dimension is not None:
                first, last = _range(dimension.get("ref", "A1"))
                present = [coordinate_key(coordinate) for coordinate in coordinates if source.cells.get((sheet.id, coordinate)) is not None]
                if present:
                    first = (min(first[0], *(row for row, _ in present)), min(first[1], *(column for _, column in present)))
                    last = (max(last[0], *(row for row, _ in present)), max(last[1], *(column for _, column in present)))
                    dimension.set("ref", self._coordinate(first) + ":" + self._coordinate(last) if first != last else self._coordinate(first))
            updates[sheet.path] = _serialize(tree, self._parts[sheet.path])
        self._recalculate(updates, removed)
        output = io.BytesIO()
        with zipfile.ZipFile(output, "w") as archive:
            archive.comment = self._comment
            for info in self._infos:
                if info.filename not in removed:
                    archive.writestr(copy.copy(info), updates.get(info.filename, self._parts[info.filename]))
        result = output.getvalue()
        actual = Workbook(result)
        expected = dict(self.cells)
        for key in changed:
            if key in source.cells:
                expected[key] = source.cells[key]
            else:
                expected.pop(key, None)
        if actual.cells != expected:
            raise WorkbookError("部分ステージ後のセル検証に失敗しました。インデックスは更新していません")
        self.assert_compatible(actual)
        return result

    @staticmethod
    def _coordinate(position):
        row, column = position
        letters = ""
        while column:
            column, remainder = divmod(column - 1, 26)
            letters = chr(65 + remainder) + letters
        return letters + str(row)

    def _replace(self, data, rows, coordinate, cell):
        row_number, column_number = coordinate_key(coordinate)
        row = rows.get(str(row_number))
        if row is None:
            if cell is None:
                return
            row = etree.Element(_tag(self._ns, "row"), r=str(row_number))
            following = next((node for node in _children(data, "row") if int(node.get("r", "0")) > row_number), None)
            rows[str(row_number)] = row
            if following is None:
                data.append(row)
            else:
                data.insert(data.index(following), row)
        nodes = _children(row, "c")
        node = next((node for node in nodes if node.get("r") == coordinate), None)
        if node is None:
            if cell is None:
                return
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
        for child in list(node):
            if _local(child) in ("f", "v", "is"):
                node.remove(child)
        node.attrib.pop("t", None)
        if cell is None:
            if set(node.attrib) == {"r"} and not len(node):
                row.remove(node)
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
