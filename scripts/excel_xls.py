"""Read BIFF8 formulas while xlrd supplies the cached cell values."""

from dataclasses import dataclass
import datetime as dt
import hashlib
import io
import json
import struct

import xlrd

from excel_formula import FormulaError, decode_formula, quote_sheet_name


EOF = 0x000A
CONTINUE = 0x003C
FORMULA = 0x0006
ARRAY = 0x0221
SHRFMLA = 0x04BC
TABLE = 0x0236
STRING = 0x0207
SUPBOOK = 0x01AE
EXTERNSHEET = 0x0017
EXTERNNAME = 0x0023
NAME = 0x0018
BOUNDSHEET = 0x0085


@dataclass(frozen=True)
class Record:
    kind: int
    start: int
    end: int
    chunks: tuple

    @property
    def data(self):
        return b"".join(self.chunks)


def records(data, start):
    """Yield a substream's records, joining following CONTINUE payloads."""
    position = start
    while position < len(data):
        record_start = position
        if len(data) - position < 4:
            raise ValueError("xls のレコードヘッダーが途中で切れています")
        kind, size = struct.unpack_from("<HH", data, position)
        position += 4
        if size > len(data) - position:
            raise ValueError("xls のレコードが途中で切れています")
        chunks = [bytes(data[position:position + size])]
        position += size
        while kind != EOF and position + 4 <= len(data):
            following, size = struct.unpack_from("<HH", data, position)
            if following != CONTINUE:
                break
            position += 4
            if size > len(data) - position:
                raise ValueError("xls の CONTINUE が途中で切れています")
            chunks.append(bytes(data[position:position + size]))
            position += size
        yield Record(kind, record_start, position, tuple(chunks))
        if kind == EOF:
            return
    raise ValueError("xls のシートに EOF がありません")


class Reader:
    """Read record chunks, retaining Unicode's continuation boundaries."""

    def __init__(self, chunks):
        self.chunks = chunks
        self.chunk = 0
        self.position = 0

    def read(self, size):
        result = bytearray()
        while size:
            if self.chunk >= len(self.chunks):
                raise ValueError("xls のレコード内のデータが不足しています")
            current = self.chunks[self.chunk]
            available = len(current) - self.position
            count = min(size, available)
            result.extend(current[self.position:self.position + count])
            self.position += count
            size -= count
            if not available:
                self.chunk += 1
                self.position = 0
        return bytes(result)

    def unpack(self, pattern):
        return struct.unpack(pattern, self.read(struct.calcsize(pattern)))

    def unicode(self, length=None, length_size=1):
        if length is None:
            length, = self.unpack("<B" if length_size == 1 else "<H")
        flags, = self.unpack("<B")
        rich_runs = self.unpack("<H")[0] if flags & 8 else 0
        extended_size = self.unpack("<I")[0] if flags & 4 else 0
        width = 2 if flags & 1 else 1
        result = bytearray()
        while length:
            if self.chunk >= len(self.chunks):
                raise ValueError("xls の文字列が途中で切れています")
            available = len(self.chunks[self.chunk]) - self.position
            if not available:
                self.chunk += 1
                self.position = 0
                # Only continued character data has a compression flag.
                flags, = self.unpack("<B")
                width = 2 if flags & 1 else 1
                continue
            count = min(length, available // width)
            if not count:
                raise ValueError("xls の Unicode 文字列の境界が不正です")
            value = self.read(count * width)
            result.extend(value if width == 2 else value.decode("latin1").encode("utf-16le"))
            length -= count
        self.read(rich_runs * 4 + extended_size)
        return result.decode("utf-16le")


def _formula_context_data(book):
    """Resolve only internal sheets, names, and this workbook's add-in names."""
    sheet_names = []
    worksheet_indexes = []
    scopes = {}
    supporting = []
    xtis = []
    names = {}
    for record in records(book.mem, book.base):
        reader = Reader(record.chunks)
        if record.kind == BOUNDSHEET:
            _, _, sheet_type = reader.unpack("<IBB")
            if sheet_type == 0:
                worksheet_indexes.append(len(sheet_names))
            sheet_names.append(reader.unicode())
        elif record.kind == SUPBOOK:
            header = reader.read(4)
            if header[2:4] == b"\x01\x04":
                kind = "self"
            elif header == b"\x01\x00\x01\x3a":
                kind = "addin"
            else:
                kind = "external"
            supporting.append((kind, []))
        elif record.kind == EXTERNSHEET:
            count, = reader.unpack("<H")
            xtis.extend(reader.unpack("<HHH") for _ in range(count))
        elif (record.kind == EXTERNNAME and supporting
              and supporting[-1][0] == "addin"):
            reader.read(6)
            supporting[-1][1].append(reader.unicode())
        elif record.kind == NAME:
            flags, _, length = reader.unpack("<HBB")
            reader.read(4)
            scope, = reader.unpack("<H")
            reader.read(4)
            name = reader.unicode(length)
            if flags & 0x20:
                name = xlrd.book.builtin_name_from_code.get(name, name)
            scopes[len(names) + 1] = scope
            names[len(names) + 1] = name
    sheets = {}
    external_names = {}
    for ixti, (support, first, last) in enumerate(xtis):
        if support >= len(supporting):
            continue
        kind, addin_names = supporting[support]
        if kind == "self" and first <= last < len(sheet_names):
            sheets[ixti] = (sheet_names[first], sheet_names[last])
        elif kind == "addin":
            for index, name in enumerate(addin_names, 1):
                external_names[ixti, index] = name
    context = {"version": 8, "sheets": sheets, "names": names,
               "external_names": external_names}
    return context, sheet_names, worksheet_indexes, scopes


def _scoped_context(data, sheet_index):
    context, sheet_names, worksheet_indexes, scopes = data
    if sheet_index is None:
        return context
    current = (worksheet_indexes[sheet_index]
               if 0 <= sheet_index < len(worksheet_indexes) else None)
    names = {}
    for index, name in context["names"].items():
        scope = scopes[index]
        if scope and scope <= len(sheet_names) and scope - 1 != current:
            name = quote_sheet_name(sheet_names[scope - 1]) + "!" + name
        names[index] = name
    return {**context, "names": names}


def formula_context(book, sheet_index=None):
    """Return decoder context, qualifying local names on other sheets."""
    return _scoped_context(_formula_context_data(book), sheet_index)


@dataclass(frozen=True)
class Formula:
    rgce: bytes
    rgcb: bytes = b""


@dataclass(frozen=True)
class FormulaRange:
    bounds: tuple
    formula: Formula
    kind: str
    payload: bytes = b""


def parsed_formula(data, offset):
    if len(data) < offset + 2:
        raise ValueError("xls の数式の長さがありません")
    size, = struct.unpack_from("<H", data, offset)
    offset += 2
    if size > len(data) - offset:
        raise ValueError("xls の数式トークンが途中で切れています")
    return Formula(data[offset:offset + size], data[offset + size:])


def sheet_formulas(book, sheet_index):
    """Return cell expressions and their shared/array/table definitions."""
    formulas = {}
    ranges = {}
    previous = None
    for record in records(book.mem, book._sh_abs_posn[sheet_index]):
        data = record.data
        if record.kind == FORMULA:
            if len(data) < 22:
                raise ValueError("xls の FORMULA が途中で切れています")
            row, col = struct.unpack_from("<HH", data)
            previous = row, col
            formulas[previous] = parsed_formula(data, 20)
        elif record.kind in (ARRAY, SHRFMLA, TABLE):
            if previous is None or len(data) < 6:
                raise ValueError("xls の数式の参照先がありません")
            bounds = struct.unpack_from("<HHBB", data)
            if record.kind == TABLE:
                if len(data) < 16:
                    raise ValueError("xls の TABLE が途中で切れています")
                formula = Formula(b"")
                kind = "table"
            else:
                offset = 12 if record.kind == ARRAY else 8
                formula = parsed_formula(data, offset)
                kind = "array" if record.kind == ARRAY else "shared"
            ranges[previous] = FormulaRange(bounds, formula, kind, data)
        # xlrd reads STRING, including its Unicode CONTINUE flags, as the
        # preceding FORMULA's cached value. Keep previous across that record.
    return formulas, ranges


def quoted(value):
    return json.dumps(value, ensure_ascii=False, allow_nan=False)


def column_letter(index):
    result = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def coordinate(row, col):
    return column_letter(col + 1) + str(row + 1)


def range_ref(bounds):
    first_row, last_row, first_col, last_col = bounds
    return coordinate(first_row, first_col) + ":" + coordinate(last_row, last_col)


def value_text(value):
    if isinstance(value, dt.datetime):
        return "datetime: " + value.isoformat()
    if isinstance(value, dt.time):
        return "time: " + value.isoformat()
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return quoted(value)


def cached_text(book, cell):
    value = cell.value
    if cell.ctype == xlrd.XL_CELL_DATE:
        value = xlrd.xldate_as_datetime(value, book.datemode)
        if 0 <= cell.value < 1:
            value = value.time()
        return value_text(value)
    if cell.ctype == xlrd.XL_CELL_BOOLEAN:
        return quoted(bool(value))
    if cell.ctype == xlrd.XL_CELL_ERROR:
        return "error: " + quoted(xlrd.error_text_from_code[value])
    return value_text(value)


def table_text(definition):
    flags, row1, col1, row2, col2 = struct.unpack_from("<5H", definition.payload, 6)
    two_variables = bool(flags & 8)
    attributes = {"t": "dataTable", "ref": range_ref(definition.bounds),
                  "dt2D": "1" if two_variables else "0",
                  "dtr": "1" if two_variables or flags & 4 else "0"}
    if not flags & 0x10:
        attributes["r1"] = coordinate(row1, col1)
    if two_variables and not flags & 0x20:
        attributes["r2"] = coordinate(row2, col2)
    if flags & 0x10:
        attributes["del1"] = "1"
    if two_variables and flags & 0x20:
        attributes["del2"] = "1"
    if flags & 1:
        attributes["ca"] = "1"
    return "data-table-formula: " + quoted(attributes)


def formula_cells(book):
    """Return each worksheet's already formatted formula overrides."""
    if book.biff_version < 80:
        return [{} for _ in range(book.nsheets)]
    context_data = _formula_context_data(book)
    result = []
    for sheet_index in range(book.nsheets):
        context = _scoped_context(context_data, sheet_index)
        formulas, ranges = sheet_formulas(book, sheet_index)
        sheet = book.sheet_by_index(sheet_index)
        output = {}
        for (row, col), formula in formulas.items():
            definition = None
            tokens = formula.rgce
            try:
                if tokens[:1] in (b"\x01", b"\x02"):
                    if len(tokens) != 5:
                        raise FormulaError("xls の配列・共有数式の参照が不正です")
                    anchor = struct.unpack_from("<HH", tokens, 1)
                    definition = ranges.get(anchor)
                    if definition is None:
                        raise FormulaError("xls の配列・共有数式の参照先がありません")
                    r1, r2, c1, c2 = definition.bounds
                    if not (r1 <= row <= r2 and c1 <= col <= c2):
                        raise FormulaError("xls の配列・共有数式の範囲が不正です")
                    if definition.kind != "shared" and (row, col) != (r1, c1):
                        continue
                    formula = definition.formula
                if definition is not None and definition.kind == "table":
                    text = table_text(definition)
                else:
                    expression = "=" + decode_formula(formula.rgce, formula.rgcb,
                                                       row=row, col=col, **context)
                    if definition is not None and definition.kind == "array":
                        text = "array-formula: " + quoted(
                            {"text": expression, "ref": range_ref(definition.bounds)})
                    else:
                        text = "formula: " + quoted(expression)
            except FormulaError:
                digest = hashlib.sha256(formula.rgce + formula.rgcb).hexdigest()[:8]
                text = ("formula-unparsed: " + quoted("#" + digest) + " "
                        + cached_text(book, sheet.cell(row, col)))
            output[row, col] = text
        result.append(output)
    return result


def render_text(data):
    book = xlrd.open_workbook(file_contents=data, on_demand=True)
    try:
        overrides = formula_cells(book)
        out = io.StringIO()
        for index, sheet in enumerate(book.sheets()):
            out.write("Sheet: " + quoted(sheet.name) + "\n")
            for row in range(sheet.nrows):
                for col, cell in enumerate(sheet.row(row)):
                    text = overrides[index].get((row, col))
                    if text is None:
                        if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                            continue
                        text = cached_text(book, cell)
                    out.write(coordinate(row, col) + "\t" + text + "\n")
        return out.getvalue()
    finally:
        book.release_resources()
