"""Lossless editing of BIFF12 cell values, formulas, and formatting.

The read-only BIFF12 model retains the original records. Selected records are
copied from the source, with target formatting, without resaving the package.
"""

import copy
import hashlib
import io
import struct
import zipfile
from array import array
from bisect import bisect_left
from collections.abc import Mapping
from functools import lru_cache
from itertools import zip_longest

import excel_stage_ooxml as ooxml
import excel_xlsb as xlsb
from excel_stage_xlsb_styles import BinaryStyles, check_styles_appended
from excel_formula import FormulaError, decode_formula


WorkbookError = ooxml.WorkbookError
Cell = ooxml.Cell
Sheet = ooxml.Sheet
_DATA_TYPES = set(range(1, 12)) | {62}
_FORMULA_TYPES = {8, 9, 10, 11}
_FAILURE = "作り直したブックの検証に失敗しました。書き込みは行っていません"


def _record(kind, payload):
    return xlsb.record_header(kind, len(payload)) + payload


def _record_positions(data):
    """Iterate record positions without allocating payloads or record objects."""
    position, length = 0, len(data)
    while position < length:
        start = position
        kind = data[position]
        position += 1
        if kind & 0x80:
            if position >= length or data[position] & 0x80:
                raise WorkbookError("xlsb のレコードの種類が不正です")
            kind = (kind & 0x7F) | (data[position] << 7)
            position += 1
        size, shift = 0, 0
        while True:
            if position >= length or shift == 28:
                raise WorkbookError("xlsb のレコードの大きさが不正です")
            value = data[position]
            position += 1
            size |= (value & 0x7F) << shift
            if not value & 0x80:
                break
            shift += 7
        end = position + size
        if end > length:
            raise WorkbookError("xlsb のレコードが途中で切れています")
        yield kind, start, position, end
        position = end


def recompute_index(data):
    """Rebuild the Excel binaryIndex offsets from uncompressed sheet bytes."""
    rows, offsets, begin, row = {}, {}, None, None
    unpack = struct.unpack_from
    for kind, start, payload, _ in _record_positions(data):
        if kind == 145:
            begin = start
        elif kind == 0:
            row = unpack("<I", data, payload)[0]
            offsets[row] = start
        elif kind in _DATA_TYPES:
            if row is None or begin is None:
                raise WorkbookError("xlsb のセルの前に行の記録がありません")
            block = unpack("<I", data, payload)[0] // 1024
            # Only the first cell in each column block participates in the index.
            rows.setdefault(row, {}).setdefault(block, start)
    starts = []
    for row in sorted(rows):
        if not starts or row >= starts[-1] + 32:
            starts.append(row)
    result = bytearray()
    if not starts:
        result += _record(42, struct.pack("<IIII", 0, 32, 0, 0) + bytes(8))
    for index, first in enumerate(starts):
        last = min(first + 32, 1048576)
        present = [row for row in range(first, last) if row in rows]
        base = begin if index == 0 else offsets[present[0]]
        row_mask, column_masks, positions = 0, [], []
        for row in present:
            row_mask |= 1 << (row - first)
            mask = 0
            for block, offset in sorted(rows[row].items()):
                mask |= 1 << block
                positions.append(offset - base)
            column_masks.append(mask)
        row_payload = struct.pack("<IQ", row_mask, base)
        row_payload += b"".join(struct.pack("<H", mask) for mask in column_masks)
        row_payload += b"".join(struct.pack("<I", offset) for offset in positions)
        unused = bytes(((last - first + 32) // 32) * 4)
        size = 16 + len(unused)
        following = len(result) + len(xlsb.record_header(42, size)) + size
        result += _record(42, struct.pack("<IIII", first, last, following, 0) + unused)
        result += _record(40, row_payload)
    result += _record(277, b"")
    return bytes(result)


def _number(value):
    # Preserve the canonical exponent notation used for large integer doubles.
    if -1e16 < value < 1e16 and float(value).is_integer():
        return str(int(value))
    return ooxml._number(repr(float(value)))


@lru_cache(maxsize=16384)
def _column_name(column):
    return xlsb._coordinate(0, column)[:-1]


def _coordinate(row, column):
    return _column_name(column) + str(row + 1)


def _context_key(context):
    return tuple(tuple(sorted((key, tuple(value) if isinstance(value, list) else value)
                              for key, value in context[name].items()))
                 for name in ("sheets", "names", "external_names"))


@lru_cache(maxsize=4096)
def _position_dependent(rgce):
    # Scan tokens, not payload bytes: strings and numeric constants may contain
    # any byte value, including the RefN/AreaN token numbers.
    position, length = 0, len(rgce)
    sizes = {0x1C: 1, 0x1D: 1, 0x1E: 2, 0x1F: 8, 0x20: 14,
             0x21: 2, 0x22: 3, 0x23: 4, 0x24: 6, 0x25: 12,
             0x26: 6, 0x27: 6, 0x28: 6, 0x29: 2, 0x2A: 6,
             0x2B: 12, 0x39: 6, 0x3A: 8, 0x3B: 14, 0x3C: 8, 0x3D: 14}
    while position < length:
        token = rgce[position]
        position += 1
        if token & 0x80:
            return True
        token = (token & 0x1F) | 0x20 if token >= 0x20 else token
        if token in (0x2C, 0x2D):
            return True
        if 3 <= token <= 0x16:
            continue
        if token == 0x17 and position + 2 <= length:
            size = 2 + 2 * struct.unpack_from("<H", rgce, position)[0]
        elif token == 0x19 and position + 3 <= length:
            flags, value = struct.unpack_from("<BH", rgce, position)
            size = 3 + ((value + 1) * 2 if flags & 4 and not flags & 0x50 else 0)
        elif token in sizes:
            size = sizes[token]
        else:
            return True
        position += size
    return position != length


@lru_cache(maxsize=65536)
def _decode_formula_cached(rgce, rgcb, context, row, col):
    return decode_formula(rgce, rgcb, version=12, row=row, col=col,
                          **dict(zip(("sheets", "names", "external_names"),
                                     (dict(values) for values in context))))


class _CellRecords(Mapping):
    """Find raw cell views from packed coordinates instead of retaining views."""
    def __init__(self, workbook):
        self.workbook = workbook

    def __getitem__(self, key):
        identifier, coordinate = key
        positions = self.workbook._cell_positions.get(identifier)
        if positions is None:
            raise KeyError(key)
        row, col = ooxml.coordinate_key(coordinate)
        value = row * 16384 + col - 1
        index = bisect_left(positions, value)
        if index == len(positions) or positions[index] != value:
            raise KeyError(key)
        return self.workbook._binary_sheets[identifier].cells[index]

    def __iter__(self):
        for identifier, item in self.workbook._binary_sheets.items():
            for raw in item.cells:
                yield identifier, _coordinate(raw.row, raw.col)

    def __len__(self):
        return sum(map(len, self.workbook._cell_positions.values()))


# Keep the name used in the fixture investigation available to callers.
recompute = recompute_index


class _BinaryWorkbook(xlsb.BinaryWorkbook):
    def _read_sheet(self, sheet, data):
        relation = self.relationships.get(sheet.rel_id)
        if relation and relation[0].endswith("/worksheet"):
            super()._read_sheet(sheet, data)


class Workbook(ooxml.Workbook):
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
            self._binary = _BinaryWorkbook(data, parts=self._parts, compact=True)
            self._load_binary()
        except WorkbookError:
            raise
        except (xlsb.WorkbookError, OSError, KeyError, RuntimeError, zipfile.BadZipFile, ValueError, struct.error) as error:
            raise WorkbookError("Excel ブックを読めません: " + str(error)) from error

    def _load_binary(self, load_cells=True):
        binary = self._binary
        self._workbook_path = binary.path
        self._workbook_rels = ooxml._rels_path(binary.path)
        self._binary_sheets, self._cell_positions = {}, {}
        self._cell_records = _CellRecords(self)
        descriptors, structure, ids, names = [], [], set(), set()
        for item in binary.sheets:
            if not item.name or item.id in ids or item.name in names:
                raise WorkbookError("シート名または sheetId が不正です")
            ids.add(item.id)
            names.add(item.name)
            relation = binary.relationships.get(item.rel_id)
            if relation is None or relation[1] is None:
                raise WorkbookError("シートの参照が不正です: " + item.name)
            kind = relation[0].rsplit("/", 1)[-1]
            structure.append((item.id, item.name, item.rel_id, item.path, kind))
            if kind != "worksheet":
                continue
            sheet = Sheet(item.id, item.name, item.path)
            descriptors.append(sheet)
            self._binary_sheets[item.id] = item
            if load_cells:
                self._load_cells(sheet, item)
        self.sheets = tuple(descriptors)
        defined = tuple((name.name, name.scope, name.rgce, name.rgcb) for name in binary.names)
        external, in_externals = [], False
        external_types = {355, 357, 358, 359, 360, 361, 362, 667} | set(range(363, 372)) | set(range(577, 589))
        workbook_data = self._parts[binary.path]
        for record in binary.workbook_records:
            if record.type == 353:
                in_externals = True
            if in_externals or record.type in external_types:
                external.append(workbook_data[record.start:record.end])
            if record.type == 354:
                in_externals = False
        relations = ooxml._xml(self._parts[self._workbook_rels], self._workbook_rels)
        for relation in relations.getroot():
            if not relation.get("Type", "").endswith("/externalLink"):
                continue
            if relation.get("TargetMode") == "External":
                raise WorkbookError("Excel の外部参照のパーツが不正です")
            path = xlsb._target(binary.path, relation.get("Target", ""))
            if path not in self._parts:
                raise WorkbookError("Excel のパーツがありません: " + path)
            external.append((relation.get("Id"), relation.get("Type"), path,
                             relation.get("TargetMode"), self._parts[path],
                             self._parts.get(ooxml._rels_path(path))))
        self._structure = (binary.path, "xlsb", tuple(structure), binary.date1904, defined, tuple(external))

    def _load_cells(self, sheet, item):
        identifier = sheet.id
        self._ranges[identifier], self._merges[identifier] = [], []
        self._row_styles[identifier], self._col_styles[identifier] = {}, []
        previous = -1
        for row in item.rows:
            if row.row <= previous or len(row.spans) > 16:
                raise WorkbookError("xlsb の行の記録が不正です")
            previous = row.row
            if row.record.payload[11] & 0x40:
                self._row_styles[identifier][row.row + 1] = row.ixfe
        for first, last, _, style, _ in item.columns:
            if not 0 <= first <= last < 16384:
                raise WorkbookError("xlsb の列の記録が不正です")
            self._col_styles[identifier].append((first + 1, last + 1, style))
        self._col_styles[identifier].sort()
        for group in (item.arrays, item.shared):
            for value in group.values():
                a, b, c, d = value.bounds
                self._ranges[identifier].append(((a + 1, c + 1), (b + 1, d + 1)))
        for bounds, _, _ in item.tables.values():
            a, b, c, d = bounds
            self._ranges[identifier].append(((a + 1, c + 1), (b + 1, d + 1)))
        for bounds, _ in item.merges:
            a, b, c, d = bounds
            self._merges[identifier].append(((a + 1, c + 1), (b + 1, d + 1)))
        context = self._binary.formula_context(item)
        prefix_starts = self._cell_prefix_starts(item)
        headers = iter(item.rows)
        header = next(headers, None)
        current_row, spans = -1, ()
        order, positions = [], array("Q")
        raw_positions = array("Q")
        context_key = _context_key(context)
        values, cell_styles, unsafe = self.cells, self.cell_styles, self._unsafe
        for row, col, style, record_type, start, kind, value, rich, rgce, rgcb, index in item.cells.iter_values():
            if row != current_row:
                while header is not None and header.row < row:
                    header = next(headers, None)
                if header is None or header.row != row:
                    raise WorkbookError("xlsb のセルの前に行の記録がありません")
                current_row, spans = row, header.spans
            if not any(first <= col <= last for first, last in spans):
                raise WorkbookError("xlsb のセルが行の列範囲の外にあります")
            key = (identifier, _coordinate(row, col))
            position = (row + 1) * 16384 + col
            raw_positions.append(position)
            if rich or record_type == 62:
                unsafe[key] = "リッチテキスト"
            cell = None
            if record_type in _FORMULA_TYPES:
                if rgce and rgce[0] == 1:
                    cell = Cell(kind, _number(value) if kind == "number" else value)
                cell = self._formula_cell(item, item.cells[index], key, cell, context, context_key, rgce, rgcb)
            elif kind != "blank":
                cell = Cell(kind, _number(value) if kind == "number" else value)
            if start in prefix_starts:
                unsafe[key] = "セルのメタデータ"
            if cell is not None:
                values[key] = cell
                order.append(key)
                positions.append(position)
            if cell is None or style != 0:
                cell_styles[key] = style
        self._cell_positions[identifier] = positions if len(positions) == len(raw_positions) else raw_positions
        self._order[identifier], self._positions[identifier] = order, positions
        self._c_counts[identifier] = len(item.cells)

    @staticmethod
    def _cell_prefix_starts(item):
        starts = {}
        records = item.records
        types = records._types if hasattr(records, "_types") else [record.type for record in records]
        if 49 not in types and 50 not in types:
            return starts
        for index, kind in enumerate(types):
            if kind not in xlsb._CELL_TYPES:
                continue
            first = index
            while first and types[first - 1] in (49, 50):
                first -= 1
            if first != index:
                starts[records[index].start] = records[first].start
        return starts

    def _formula_cell(self, sheet, raw, key, cached, context, context_key=None, rgce=None, rgcb=None):
        if rgce is None:
            rgce, rgcb = raw.rgce, raw.rgcb
        try:
            if not rgce:
                raise FormulaError("数式のトークン列が空です")
            if rgce[0] == 1:
                if len(rgce) != 5 or len(rgcb) < 4:
                    raise FormulaError("配列・共有数式の参照が壊れています")
                anchor = (struct.unpack("<I", rgce[1:])[0], struct.unpack("<I", rgcb[:4])[0])
                group = sheet.arrays.get(anchor) or sheet.shared.get(anchor)
                if anchor in sheet.tables:
                    if (raw.row, raw.col) == anchor:
                        self._unsafe[key] = "数式の拡張機能"
                        return Cell("formula", "=")
                    return cached
                if group is None:
                    raise FormulaError("配列・共有数式の式が見つかりません")
                if anchor in sheet.arrays and (raw.row, raw.col) != anchor:
                    return cached
                rgce, rgcb = group.rgce, group.rgcb
                if (raw.row, raw.col) == anchor:
                    self._unsafe[key] = "数式の拡張機能"
            context_key = _context_key(context) if context_key is None else context_key
            if len(rgce) + len(rgcb) <= 1024:
                row, col = (raw.row, raw.col) if _position_dependent(rgce) else (0, 0)
                formula = _decode_formula_cached(rgce, rgcb, context_key, row, col)
            else:
                formula = decode_formula(rgce, rgcb, version=12, row=raw.row, col=raw.col, **context)
            return Cell("formula", "=" + formula)
        except FormulaError:
            digest = hashlib.sha256(rgce + rgcb).hexdigest()[:8]
            self._unsafe[key] = "式に戻せない数式"
            return Cell("formula", "=（式に戻せない数式 " + digest + "）")

    def stylesheet(self):
        if self._stylesheet is None:
            paths = [path for kind, path in self._binary.relationships.values()
                     if kind.endswith("/styles")]
            if len(paths) > 1:
                raise WorkbookError("Excel の書式の参照が重複しています")
            self._styles_path = paths[0] if paths else None
            self._stylesheet = BinaryStyles(self._parts.get(self._styles_path), self._styles_path)
            self._stylesheet.parts = self._parts
        return self._stylesheet

    def compare_formats(self, other):
        return ooxml.FormatComparison(self, other)

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
            ooxml.coordinate_key(key[1])
            reason = comparison.unsupported_reason(key)
            if reason:
                raise WorkbookError(sheets[key[0]].name + "!" + key[1] + " の書式は反映できません: " + reason)
        restyled = {key for key in formats if comparison.differs(key)}
        if not changed and not restyled:
            return self.data
        self._index_cache = {}
        updates, removed, styles = {}, set(), {}
        if restyled:
            importer = comparison.target_styles.importer(comparison.source_styles)
            for key in sorted(restyled):
                styles[key] = importer.index_for(source.effective_style(key))
            if importer.modified:
                updates[self._styles_path] = importer.serialize(self._parts[self._styles_path])
        for sheet in self.sheets:
            coordinates = {coordinate for identifier, coordinate in changed if identifier == sheet.id}
            sheet_styles = {coordinate: index for (identifier, coordinate), index in styles.items()
                            if identifier == sheet.id}
            if coordinates or sheet_styles:
                if sheet_styles:
                    updates[sheet.path] = self._rebuild_sheet(sheet, source, coordinates, sheet_styles)
                else:
                    updates[sheet.path] = self._rebuild_sheet(sheet, source, coordinates)
                self._reindex(sheet, updates, removed)
        if changed:
            self._recalculate(updates, removed)
        self._remove_content_types(updates, removed)
        output = io.BytesIO()
        with zipfile.ZipFile(output, "w") as archive:
            archive.comment = self._comment
            for info in self._infos:
                if info.filename not in removed:
                    archive.writestr(copy.copy(info), updates.get(info.filename, self._parts[info.filename]))
        result = output.getvalue()
        try:
            self._validate_result(result, source, changed, updates, removed, styles)
        finally:
            self._index_cache.clear()
        return result

    def _replacement(self, source, key, style=None):
        original = self._cell_records.get(key)
        incoming = source._cell_records.get(key)
        if style is None:
            style = self.effective_style(key)
        flags = struct.unpack_from("<I", original.record.payload, 4)[0] & 0xFF000000 if original else 0
        row, column = ooxml.coordinate_key(key[1])
        head = struct.pack("<II", column - 1, flags | style)
        if key not in source.cells:
            inherited = self._inherited_style(key[0], row, column)
            if original is None and style == inherited:
                return b""
            return _record(1, head) if style or inherited else b""
        if incoming is None:
            raise WorkbookError("取り込み元のセルの記録がありません")
        if incoming.record.type == 7 and source is not self:
            value = incoming.value.encode("utf-16le")
            return _record(6, head + struct.pack("<I", len(value) // 2) + value)
        sheet = source._binary_sheets[key[0]]
        original = source._parts[sheet.path]
        header = original[incoming.record.start:incoming.record.payload_start]
        return header + head + incoming.record.payload[8:]

    @staticmethod
    def _expand_spans(spans, columns):
        columns = tuple(columns)
        if all(any(first <= column <= last for first, last in spans) for column in columns):
            return tuple(spans)
        # Each 1024-column block needs one encompassing range. Coalesce all
        # original ranges in that block before expanding it for a new cell.
        blocks = {}
        for first, last in spans:
            for block in range(first // 1024, last // 1024 + 1):
                a, b = max(first, block * 1024), min(last, block * 1024 + 1023)
                old = blocks.get(block, (a, b))
                blocks[block] = (min(old[0], a), max(old[1], b))
        for column in columns:
            block = column // 1024
            first, last = blocks.get(block, (column, column))
            blocks[block] = (min(first, column), max(last, column))
        return tuple(blocks[block] for block in sorted(blocks))

    @staticmethod
    def _dimension_position(records):
        if not records or records[0].type != 129:
            raise WorkbookError("xlsb のシートの始まりがありません")
        previous = records[1] if len(records) > 1 and records[1].type == 147 else records[0]
        return previous.end

    @staticmethod
    def _calculation_position(records):
        # Workbook ABNF places CalcProp after BUNDLESHS, optional FNGROUP,
        # optional EXTERNALS, and all names, before OleSize and later parts.
        records = [record for record in records if record.type != 157]
        index = next((index + 1 for index, record in enumerate(records) if record.type == 144), None)
        if index is None:
            raise WorkbookError("xlsb のシートの一覧の終わりがありません")
        for begin, end in ((664, 666), (353, 354)):
            if index < len(records) and records[index].type == begin:
                index = next((last + 1 for last in range(index + 1, len(records))
                              if records[last].type == end), None)
                if index is None:
                    raise WorkbookError("xlsb のブックの記録が不正です")
        while index < len(records) and records[index].type == 39:
            index += 1
        if index >= len(records):
            raise WorkbookError("xlsb のブックの終わりがありません")
        return records[index].start

    def _rebuild_sheet(self, sheet, source, coordinates, styles=None):
        styles = styles or {}
        selected = set(coordinates) | set(styles)
        item = self._binary_sheets[sheet.id]
        original = self._parts[sheet.path]
        replacements, new_rows, row_columns = {}, {}, {}
        old_rows = {row.row: row for row in item.rows}
        cells_by_row = {}
        prefix_starts = self._cell_prefix_starts(item)
        record_types = item.records._types
        written = {}
        for coordinate in selected:
            key = (sheet.id, coordinate)
            row, column = ooxml.coordinate_key(coordinate)
            row, column = row - 1, column - 1
            owner = source if coordinate in coordinates else self
            value = self._replacement(owner, key, styles.get(coordinate))
            written[coordinate] = value
            old = self._cell_records.get(key)
            if old:
                replacements[old.start] = (old.end, value)
            elif value:
                new_rows.setdefault(row, {})[column] = value
            if value:
                row_columns.setdefault(row, set()).add(column)
        for row, columns in row_columns.items():
            if row in old_rows:
                header = old_rows[row]
                spans = self._expand_spans(header.spans, columns)
                if spans != header.spans:
                    payload = header.record.payload[:13] + struct.pack("<I", len(spans))
                    payload += b"".join(struct.pack("<II", first, last) for first, last in spans)
                    payload += header.record.payload[17 + len(header.spans) * 8:]
                    replacements[header.record.start] = (header.record.end, _record(0, payload))
        # Insert missing cells before the next existing cell, retaining every
        # intervening formula/extension record in its original position.
        inserts = {}
        for row in sorted(set(new_rows) & set(old_rows)):
            positions = self._cell_positions[sheet.id]
            first = bisect_left(positions, (row + 1) * 16384)
            last = bisect_left(positions, (row + 2) * 16384)
            cells_by_row[row] = [item.cells[index] for index in range(first, last)]
            for column, value in sorted(new_rows[row].items()):
                following = next((prefix_starts.get(cell.start, cell.start) for cell in cells_by_row.get(row, [])
                                  if cell.col > column), None)
                if following is None:
                    cells = cells_by_row.get(row, [])
                    following = cells[-1].end if cells else old_rows[row].record.end
                    for index, kind in enumerate(record_types):
                        if kind in (426, 427, 428):
                            attached = item.records[index]
                            if attached.start == following:
                                following = attached.end
                inserts.setdefault(following, []).append(value)
        height = 300
        for index, kind in enumerate(record_types):
            if kind == 485:
                record = item.records[index]
                if len(record.payload) < 8:
                    raise WorkbookError("xlsb の既定の行の高さが不正です")
                height = struct.unpack_from("<H", record.payload, 6)[0]
        row_indices = ({struct.unpack_from("<I", item.records[index].payload)[0]: index
                        for index, kind in enumerate(record_types) if kind == 0}
                       if set(new_rows) - set(old_rows) else {})
        end = next((item.records[index].start for index, kind in enumerate(record_types) if kind == 146), None)
        if end is None:
            raise WorkbookError("xlsb のシートのデータの終わりがありません")
        for row in sorted(set(new_rows) - set(old_rows)):
            spans = self._expand_spans((), new_rows[row])
            payload = struct.pack("<IIHBBBI", row, 0, height, 0, 0, 0, len(spans))
            payload += b"".join(struct.pack("<II", first, last) for first, last in spans)
            value = _record(0, payload) + b"".join(value for _, value in sorted(new_rows[row].items()))
            following_rows = [number for number in old_rows if number > row]
            if following_rows:
                index = row_indices[min(following_rows)]
                # RwDescent, often wrapped in AC records, belongs to the
                # following row and must remain beside that original header.
                while index > 0:
                    if item.records[index - 1].type == 1024:
                        index -= 1
                    elif item.records[index - 1].type == 38:
                        start, depth = index - 1, 1
                        while start and depth:
                            start -= 1
                            depth += (item.records[start].type == 38) - (item.records[start].type == 37)
                        if depth or not any(record.type == 1024 for record in item.records[start:index]):
                            break
                        index = start
                    else:
                        break
                position = item.records[index].start
            else:
                position = end
            inserts.setdefault(position, []).append(value)
        present = [ooxml.coordinate_key(coordinate) for coordinate, value in written.items() if value]
        for index, kind in enumerate(record_types):
            if kind == 148 and present:
                record = item.records[index]
                a, b, c, d = struct.unpack("<IIII", record.payload)
                a, b = min(a, *(row - 1 for row, _ in present)), max(b, *(row - 1 for row, _ in present))
                c, d = min(c, *(col - 1 for _, col in present)), max(d, *(col - 1 for _, col in present))
                replacements[record.start] = (record.end, _record(148, struct.pack("<IIII", a, b, c, d)))
        if present and 148 not in record_types:
            bounds = ([cell.row for cell in item.cells] + [row - 1 for row, _ in present],
                      [cell.col for cell in item.cells] + [col - 1 for _, col in present])
            begin = self._dimension_position(item.records)
            dimension = struct.pack("<IIII", min(bounds[0]), max(bounds[0]), min(bounds[1]), max(bounds[1]))
            inserts.setdefault(begin, []).insert(0, _record(148, dimension))
        # Copy unchanged stretches in one piece. A sparse edit must not create
        # a temporary object for every cell or split every record into bytes.
        result, cursor = bytearray(), 0
        for position in sorted(set(replacements) | set(inserts)):
            if position < cursor:
                raise WorkbookError("xlsb のセルの挿入位置が不正です")
            result += original[cursor:position]
            result += b"".join(inserts.get(position, ()))
            if position in replacements:
                cursor, value = replacements[position]
                result += value
            else:
                cursor = position
        result += original[cursor:]
        return bytes(result)

    def _remove_relationships(self, owner, suffix, updates, removed):
        path = ooxml._rels_path(owner)
        if path not in self._parts:
            return []
        tree = ooxml._xml(updates.get(path, self._parts[path]), path)
        found = []
        for node in list(tree.getroot()):
            if node.get("Type", "").endswith(suffix):
                if node.get("TargetMode") == "External":
                    raise WorkbookError("Excel の内部の参照が不正です")
                target = xlsb._target(owner, node.get("Target", ""))
                found.append(target)
                tree.getroot().remove(node)
                removed.add(target)
                removed.add(ooxml._rels_path(target))
        if found:
            updates[path] = ooxml._serialize(tree, self._parts[path])
        return found

    def _index_for(self, data):
        cache = getattr(self, "_index_cache", None)
        if cache is None:
            self._index_cache = cache = {}
        key = id(data)
        if key not in cache:
            cache[key] = (data, recompute_index(data))
        return cache[key][1]

    def _reindex(self, sheet, updates, removed):
        path = ooxml._rels_path(sheet.path)
        if path not in self._parts:
            return
        tree = ooxml._xml(self._parts[path], path)
        indexes = [node for node in tree.getroot() if node.get("Type", "").endswith("/xlBinaryIndex")]
        if not indexes:
            return
        targets = [xlsb._target(sheet.path, node.get("Target", "")) for node in indexes]
        if len(targets) == 1 and self._parts.get(targets[0]) == self._index_for(self._parts[sheet.path]):
            updates[targets[0]] = self._index_for(updates[sheet.path])
        else:
            self._remove_relationships(sheet.path, "/xlBinaryIndex", updates, removed)

    def _remove_content_types(self, updates, removed):
        path = "[Content_Types].xml"
        tree = ooxml._xml(updates.get(path, self._parts[path]), path)
        changed = False
        for node in list(tree.getroot()):
            if node.get("PartName", "").lstrip("/") in removed:
                tree.getroot().remove(node)
                changed = True
        if changed:
            updates[path] = ooxml._serialize(tree, self._parts[path])

    def _recalculate(self, updates, removed):
        original = self._parts[self._workbook_path]
        record = self._binary.calc_record
        if record:
            if len(record.payload) < 26:
                raise WorkbookError("xlsb の再計算の設定が不正です")
            payload = bytearray(record.payload)
            payload[24] |= 1
            updates[self._workbook_path] = original[:record.payload_start] + payload + original[record.end:]
        else:
            # CalcId, calcMode, cCalcCount, relative tolerance, threading count,
            # and flags. Excel defaults to automatic calculation.
            payload = struct.pack("<IIIdIH", 0, 1, 100, 0.001, 0, 1)
            end = self._calculation_position(self._binary.workbook_records)
            updates[self._workbook_path] = original[:end] + _record(157, payload) + original[end:]
        self._remove_relationships(self._workbook_path, "/calcChain", updates, removed)

    def _validate_result(self, result, source, changed, updates, removed, styles=None):
        styles = styles or {}
        touched = changed | set(styles)
        touched_sheets = {identifier for identifier, _ in touched}
        try:
            expected_parts = {path: data for path, data in self._parts.items() if path not in removed}
            expected_parts.update(updates)
            with zipfile.ZipFile(io.BytesIO(result)) as archive:
                names = archive.namelist()
                expected_names = [info.filename for info in self._infos if info.filename not in removed]
                if names != expected_names or archive.comment != self._comment:
                    raise WorkbookError(_FAILURE)
                if any(archive.read(name) != expected_parts[name] for name in names):
                    raise WorkbookError(_FAILURE)
            # The ZIP has been checked byte for byte. Reuse those immutable part
            # bytes and parse only changed sheets; untouched cells are unchanged.
            actual = Workbook.__new__(Workbook)
            actual.data = result
            actual._reset()
            actual._parts = expected_parts
            actual._binary = _BinaryWorkbook(result, parts=expected_parts,
                                            sheet_paths={sheet.path for sheet in self.sheets
                                                         if sheet.id in touched_sheets},
                                            compact=True)
            actual._load_binary(load_cells=False)
            actual._cell_records = {}
            touched_positions = {}
            for key in touched:
                row, col = ooxml.coordinate_key(key[1])
                touched_positions.setdefault(key[0], {})[(row - 1) * 16384 + col - 1] = key
            for sheet in actual.sheets:
                if sheet.id not in touched_sheets:
                    continue
                positions = touched_positions[sheet.id]
                cells = actual._binary_sheets[sheet.id].cells
                for position, key in positions.items():
                    row, col = divmod(position, 16384)
                    first = bisect_left(cells._rows, row)
                    last = bisect_left(cells._rows, row + 1)
                    for index in range(first, last):
                        raw = cells[index]
                        if raw.col == col:
                            actual._cell_records[key] = raw
                            break
            self._validate_parts(actual, changed, removed, set(styles))
            self.assert_compatible(actual)
            if actual._binary.strings != self._binary.strings:
                raise WorkbookError(_FAILURE)
            if changed:
                if not actual._binary.calc_record or not actual._binary.calc_record.payload[24] & 1:
                    raise WorkbookError(_FAILURE)
                if any(kind.endswith("/calcChain") for kind, _ in actual._binary.relationships.values()):
                    raise WorkbookError(_FAILURE)
            contexts = {}
            for key in touched:
                owner = source if key in changed else self
                row, column = ooxml.coordinate_key(key[1])
                inherited = self._inherited_style(key[0], row, column)
                style = styles.get(key, self.effective_style(key))
                before = self._cell_records.get(key)
                expected_present = key in owner.cells or (bool(style or inherited) and
                                                          not (before is None and style == inherited))
                if (key in actual._cell_records) != expected_present:
                    raise WorkbookError(_FAILURE)
                if key in styles:
                    if actual.stylesheet().exact(style) != source.stylesheet().exact(source.effective_style(key)):
                        raise WorkbookError(_FAILURE)
                if not expected_present:
                    continue
                written = actual._cell_records[key]
                if key in changed:
                    kind, value, _, rgce, rgcb = written.read()
                    parsed = None if kind == "blank" else Cell(kind, _number(value) if kind == "number" else value)
                    if written.record.type in _FORMULA_TYPES:
                        item = actual._binary_sheets[key[0]]
                        if key[0] not in contexts:
                            context = actual._binary.formula_context(item)
                            contexts[key[0]] = context, _context_key(context)
                        context, context_key = contexts[key[0]]
                        parsed = actual._formula_cell(item, written, key, parsed, context, context_key, rgce, rgcb)
                    if parsed != source.cells.get(key):
                        raise WorkbookError(_FAILURE)
                if written.ixfe != style:
                    raise WorkbookError(_FAILURE)
                old_flags = struct.unpack_from("<I", before.record.payload, 4)[0] & 0xFF000000 if before else 0
                new_flags = struct.unpack_from("<I", written.record.payload, 4)[0] & 0xFF000000
                if old_flags != new_flags:
                    raise WorkbookError(_FAILURE)
                head = struct.pack("<II", column - 1, old_flags | style)
                incoming = owner._cell_records.get(key)
                if key not in owner.cells:
                    expected_record = xlsb.record_header(1, 8) + head
                elif incoming.record.type == 7 and key in changed:
                    value = incoming.value.encode("utf-16le")
                    payload = head + struct.pack("<I", len(value) // 2) + value
                    expected_record = xlsb.record_header(6, len(payload)) + payload
                else:
                    source_data = owner._parts[owner._binary_sheets[key[0]].path]
                    expected_record = source_data[incoming.start:incoming.record.payload_start]
                    expected_record += head + source_data[incoming.record.payload_start + 8:incoming.end]
                written_data = actual._parts[actual._binary_sheets[key[0]].path]
                if written_data[written.start:written.end] != expected_record:
                    raise WorkbookError(_FAILURE)
            # Every unselected cell record and all its metadata remain bytewise
            # identical; selected records above match the source independently,
            # including cached values, formula tokens and their extra data.
            for sheet in self.sheets:
                if sheet.id not in touched_sheets:
                    continue
                self._validate_sheet(sheet, actual, touched)
                path = ooxml._rels_path(sheet.path)
                if path in actual._parts:
                    rels = ooxml._xml(actual._parts[path], path)
                    for node in rels.getroot():
                        if node.get("Type", "").endswith("/xlBinaryIndex"):
                            index = xlsb._target(sheet.path, node.get("Target", ""))
                            if actual._parts.get(index) != self._index_for(actual._parts[sheet.path]):
                                raise WorkbookError(_FAILURE)
        except (WorkbookError, xlsb.WorkbookError, OSError, ValueError, KeyError, struct.error, zipfile.BadZipFile) as error:
            raise WorkbookError(_FAILURE) from error

    def _validate_parts(self, actual, changed, removed, restyled=()):
        touched = {identifier for identifier, _ in changed | set(restyled)}
        expected_removed = set()
        for owner, suffix in ([(self._workbook_path, "/calcChain")] if changed else []) + [
                (sheet.path, "/xlBinaryIndex") for sheet in self.sheets if sheet.id in touched]:
            path = ooxml._rels_path(owner)
            if path not in self._parts:
                continue
            tree = ooxml._xml(self._parts[path], path)
            matches = [node for node in tree.getroot() if node.get("Type", "").endswith(suffix)]
            targets = [xlsb._target(owner, node.get("Target", "")) for node in matches]
            if suffix == "/xlBinaryIndex" and len(targets) == 1:
                if self._parts.get(targets[0]) == self._index_for(self._parts[owner]):
                    continue
            for target in targets:
                expected_removed.add(target)
                expected_removed.add(ooxml._rels_path(target))
        if removed != expected_removed:
            raise WorkbookError(_FAILURE)
        allowed = {self._workbook_path, self._workbook_rels, "[Content_Types].xml"}
        for sheet in self.sheets:
            if sheet.id not in touched:
                continue
            allowed.add(sheet.path)
            path = ooxml._rels_path(sheet.path)
            allowed.add(path)
            if path in self._parts:
                tree = ooxml._xml(self._parts[path], path)
                for node in tree.getroot():
                    if node.get("Type", "").endswith("/xlBinaryIndex"):
                        allowed.add(xlsb._target(sheet.path, node.get("Target", "")))
        if restyled:
            self.stylesheet()
            allowed.add(self._styles_path)
            check_styles_appended(self._parts[self._styles_path], actual._parts[self._styles_path])
        for path, content in actual._parts.items():
            if path not in allowed and content != self._parts.get(path):
                raise WorkbookError(_FAILURE)
        before = self._parts[self._workbook_path]
        after = actual._parts[self._workbook_path]
        old_records, new_records = self._binary.workbook_records, actual._binary.workbook_records
        if ([before[record.start:record.end] for record in old_records if record.type != 157] !=
                [after[record.start:record.end] for record in new_records if record.type != 157]):
            raise WorkbookError(_FAILURE)
        if not changed:
            if before != after:
                raise WorkbookError(_FAILURE)
        else:
            if sum(record.type == 157 for record in new_records) != 1:
                raise WorkbookError(_FAILURE)
            old_calc, new_calc = self._binary.calc_record, actual._binary.calc_record
            if old_calc:
                expected = bytearray(old_calc.payload)
                expected[24] |= 1
                expected = before[old_calc.start:old_calc.payload_start] + expected
            else:
                expected = _record(157, struct.pack("<IIIdIH", 0, 1, 100, 0.001, 0, 1))
            if after[new_calc.start:new_calc.end] != expected:
                raise WorkbookError(_FAILURE)
            next_position = self._calculation_position(new_records)
            if new_calc.end != next_position:
                raise WorkbookError(_FAILURE)
        # Relationship/content-type edits may remove only the rows belonging
        # to removed binary indexes and the calculation chain.
        for path in allowed:
            if path not in self._parts or path not in actual._parts:
                continue
            if path != "[Content_Types].xml" and not path.endswith(".rels"):
                continue
            old_tree = ooxml._xml(self._parts[path], path)
            new_tree = ooxml._xml(actual._parts[path], path)
            old_root, new_root = old_tree.getroot(), new_tree.getroot()
            if (old_root.tag, dict(old_root.attrib), old_root.text, old_root.tail) != (
                    new_root.tag, dict(new_root.attrib), new_root.text, new_root.tail):
                raise WorkbookError(_FAILURE)
            if path == "[Content_Types].xml":
                survivors = [node for node in old_tree.getroot()
                             if node.get("PartName", "").lstrip("/") not in removed]
            else:
                owner = self._workbook_path if path == self._workbook_rels else next(
                    sheet.path for sheet in self.sheets if ooxml._rels_path(sheet.path) == path)
                survivors = [node for node in old_tree.getroot()
                             if (node.get("TargetMode") == "External" or
                                 xlsb._target(owner, node.get("Target", "")) not in removed)]
            if [ooxml._raw(node) for node in survivors] != [ooxml._raw(node) for node in new_tree.getroot()]:
                raise WorkbookError(_FAILURE)

    def _validate_sheet(self, sheet, actual, changed):
        old, new = self._binary_sheets[sheet.id], actual._binary_sheets[sheet.id]
        old_data, new_data = self._parts[sheet.path], actual._parts[sheet.path]
        old_rows, new_rows = {row.row: row for row in old.rows}, {row.row: row for row in new.rows}
        previous = -1
        for row in new.rows:
            if row.row <= previous:
                raise WorkbookError(_FAILURE)
            previous = row.row
        if not old_rows.keys() <= new_rows.keys():
            raise WorkbookError(_FAILURE)
        touched_rows = {ooxml.coordinate_key(coordinate)[0] - 1 for identifier, coordinate in changed if identifier == sheet.id}
        for row, before in old_rows.items():
            after = new_rows[row]
            before_record, after_record = before.record, after.record
            if old_data[before_record.start:before_record.end] == new_data[after_record.start:after_record.end]:
                continue
            before_payload, after_payload = before_record.payload, after_record.payload
            before_spans, after_spans = before.spans, after.spans
            if (row not in touched_rows or len(after_spans) > 16 or
                    before_payload[:13] != after_payload[:13] or
                    before_payload[17 + 8 * len(before_spans):] != after_payload[17 + 8 * len(after_spans):]):
                raise WorkbookError(_FAILURE)
            for first, last in before_spans:
                if not any(a <= first <= last <= b for a, b in after_spans):
                    raise WorkbookError(_FAILURE)
        new_numbers = new_rows.keys() - old_rows.keys()
        if not new_numbers <= touched_rows:
            raise WorkbookError(_FAILURE)
        height = next((struct.unpack_from("<H", record.payload, 6)[0]
                       for record in old.records if record.type == 485), 300)
        for row in new_numbers:
            header = new_rows[row]
            if len(header.spans) > 16:
                raise WorkbookError(_FAILURE)
            expected = struct.pack("<IIHBBBI", row, 0, height, 0, 0, 0, len(header.spans))
            expected += b"".join(struct.pack("<II", first, last) for first, last in header.spans)
            if new_data[header.record.start:header.record.end] != _record(0, expected):
                raise WorkbookError(_FAILURE)
        chosen = {((ooxml.coordinate_key(coordinate)[0] - 1) * 16384 +
                   ooxml.coordinate_key(coordinate)[1] - 1)
                  for identifier, coordinate in changed if identifier == sheet.id}
        def protected(item, data, check=False):
            row, last_cell, spans = None, None, ()
            prefixes, prefix_records, metadata_owners = {}, set(), {}
            records = item.records
            types = records._types if hasattr(records, "_types") else [record.type for record in records]
            if 49 in types or 50 in types:
                for index, kind in enumerate(types):
                    if kind != 0:
                        continue
                    cell_row = struct.unpack_from("<I", records[index].payload)[0]
                    following = index + 1
                    while following < len(records) and types[following] not in (0, 146):
                        if types[following] in xlsb._CELL_TYPES:
                            cell_col = struct.unpack_from("<I", records[following].payload)[0]
                            previous = following - 1
                            while previous > index and types[previous] in (49, 50):
                                metadata_owners[records[previous].start] = cell_row * 16384 + cell_col
                                previous -= 1
                        following += 1
            for index, kind in enumerate(types):
                if kind != 0:
                    continue
                first = index
                while first > 0:
                    if types[first - 1] == 1024:
                        first -= 1
                    elif types[first - 1] == 38:
                        start, depth = first - 1, 1
                        while start and depth:
                            start -= 1
                            depth += (types[start] == 38) - (types[start] == 37)
                        if depth or 1024 not in types[start:first]:
                            break
                        first = start
                    else:
                        break
                if first != index:
                    prefixes[records[first].start] = struct.unpack_from("<I", records[index].payload)[0]
                    prefix_records.update(records[position].start for position in range(first, index))
            for kind, start, payload, end in _record_positions(data):
                if start in prefixes:
                    row, last_cell = prefixes[start], None
                if kind == 146:
                    row, last_cell = None, None
                if kind == 0:
                    row = struct.unpack_from("<I", data, payload)[0]
                    if check:
                        spans = new_rows[row].spans
                elif kind in xlsb._CELL_TYPES:
                    col = struct.unpack_from("<I", data, payload)[0]
                    key = row * 16384 + col
                    last_cell = key
                    if check:
                        if not any(first <= col <= last for first, last in spans):
                            raise WorkbookError(_FAILURE)
                        if new.dimensions is not None:
                            a, b, c, d = new.dimensions
                            if not (a <= row <= b and c <= col <= d):
                                raise WorkbookError(_FAILURE)
                    if key not in chosen:
                        yield ("cell", key, data[start:end])
                elif kind != 148:
                    anchor = last_cell if kind in (426, 427, 428) else metadata_owners.get(start)
                    yield ("record", row, start in prefix_records, anchor, kind, data[start:end])
        sentinel = object()
        if any(before != after for before, after in
               zip_longest(protected(old, old_data), protected(new, new_data, check=True), fillvalue=sentinel)):
            raise WorkbookError(_FAILURE)
        old_dimension, new_dimension = old.dimensions, new.dimensions
        if old_dimension is not None:
            if new_dimension is None:
                raise WorkbookError(_FAILURE)
            a, b, c, d = old_dimension
            e, f, g, h = new_dimension
            if not (e <= a <= b <= f and g <= c <= d <= h):
                raise WorkbookError(_FAILURE)
        dimensions = [record for record in new.records if record.type == 148]
        if len(dimensions) > 1:
            raise WorkbookError(_FAILURE)
        if new_dimension is not None:
            if dimensions[0].start != self._dimension_position(new.records):
                raise WorkbookError(_FAILURE)
