"""Independently encoded BIFF12 records for reader boundary tests."""

import io
import struct
import zipfile


MAIN_TYPE = "application/vnd.ms-excel.sheet.binary.macroEnabled.main"


def record(kind, payload=b""):
    if not 0 <= kind <= 0x3FFF or not 0 <= len(payload) <= 0x0FFFFFFF:
        raise ValueError("record out of range")
    header = bytes([kind]) if kind < 128 else bytes([(kind & 127) | 128, kind >> 7])
    size = len(payload)
    while size >= 128:
        header += bytes([(size & 127) | 128])
        size >>= 7
    return header + bytes([size]) + payload


def wide(value):
    if value is None:
        return struct.pack("<I", 0xFFFFFFFF)
    data = value.encode("utf-16le")
    return struct.pack("<I", len(data) // 2) + data


def parsed_formula(rgce, rgcb=b""):
    return struct.pack("<I", len(rgce)) + rgce + struct.pack("<I", len(rgcb)) + rgcb


def ptg_int(value):
    return b"\x1e" + struct.pack("<H", value)


def ptg_num(value):
    return b"\x1f" + struct.pack("<d", value)


def ptg_string(value):
    data = value.encode("utf-16le")
    return b"\x17" + struct.pack("<H", len(data) // 2) + data


def ptg_ref(row, col, *, relative_row=True, relative_col=True, shared=False):
    flags = (col & 0x3FFF) | (0x8000 if relative_row else 0) | (0x4000 if relative_col else 0)
    return bytes([0x4C if shared else 0x44]) + struct.pack("<IH", row & 0xFFFFFFFF, flags)


def ptg_exp(row, col):
    return b"\x01" + struct.pack("<I", row), struct.pack("<I", col)


def row_header(row, ixfe=0, spans=(), *, height=300, ghost=False):
    payload = struct.pack("<IIHBBBI", row, ixfe, height, 0, 0x40 if ghost else 0, 0, len(spans))
    return record(0, payload + b"".join(struct.pack("<II", *span) for span in spans))


def cell(col, kind, value=None, *, ixfe=0, rgce=None, rgcb=b"", phonetic=False):
    header = struct.pack("<II", col, ixfe | (0x01000000 if phonetic else 0))
    if kind == "blank":
        return record(1, header)
    if kind == "rk":
        return record(2, header + struct.pack("<I", value))
    types = {"number": 5, "string": 6, "boolean": 4, "error": 3, "shared-string": 7, "rich-string": 62}
    formula_types = {"number": 9, "string": 8, "boolean": 10, "error": 11}
    kind_id = formula_types[kind] if rgce is not None else types[kind]
    if kind == "number":
        payload = struct.pack("<d", value)
    elif kind in ("boolean", "error"):
        payload = bytes([value])
    elif kind == "shared-string":
        payload = struct.pack("<I", value)
    elif kind == "rich-string":
        payload = value
    else:
        payload = wide(value)
    if rgce is not None:
        payload += b"\0\0" + parsed_formula(rgce, rgcb)
    return record(kind_id, header + payload)


def formula_range(bounds, rgce, rgcb=b"", *, array=False):
    return record(426 if array else 427, struct.pack("<IIII", *bounds) + (b"\0" if array else b"") + parsed_formula(rgce, rgcb))


def rich_string(value, *, runs=(), phonetic=None, phonetic_runs=(),
                phonetic_font=1, phonetic_flags=0x0035):
    flags = bool(runs) | (2 if phonetic is not None else 0)
    payload = bytes([flags]) + wide(value)
    if runs:
        payload += struct.pack("<I", len(runs)) + b"".join(struct.pack("<HH", *run) for run in runs)
    if phonetic is not None:
        payload += wide(phonetic) + struct.pack("<I", len(phonetic_runs))
        payload += b"".join(struct.pack("<HHH", *run) for run in phonetic_runs)
        payload += struct.pack("<HH", phonetic_font, phonetic_flags)
    return payload


def worksheet(rows, *, dimensions=None, default_height=None, columns=(), before_data=b"", after_data=b""):
    prefix = record(129)
    if dimensions is not None:
        prefix += record(148, struct.pack("<IIII", *dimensions))
    if default_height is not None:
        prefix += record(485, struct.pack("<IHHI", 0xFFFFFFFF, 8, default_height, 0))
    for first, last, ixfe in columns:
        prefix += record(60, struct.pack("<IIIIH", first, last, 2560, ixfe, 0))
    return prefix + before_data + record(145) + b"".join(rows) + record(146) + after_data + record(130)


def package(sheets, *, shared_strings=(), formats=None, cell_formats=(0,), date1904=False, workbook_records=b""):
    """sheets is a sequence of (name, encoded worksheet records) pairs."""
    book = record(131) + record(153, struct.pack("<III", bool(date1904), 0, 0)) + record(143)
    rels = []
    for index, (name, _) in enumerate(sheets, 1):
        book += record(156, struct.pack("<II", 0, index) + wide("rId" + str(index)) + wide(name))
        rels.append(f'<Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{index}.bin"/>')
    book += record(144) + workbook_records + record(132)
    styles = record(278)
    for index, code in (formats or {}).items():
        styles += record(44, struct.pack("<H", index) + wide(code))
    styles += record(617, struct.pack("<I", len(cell_formats)))
    for number_format in cell_formats:
        styles += record(47, struct.pack("<HH", 0, number_format) + bytes(12))
    styles += record(618) + record(279)
    rels.append('<Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.bin"/>')
    strings = record(159, struct.pack("<II", len(shared_strings), len(shared_strings)))
    for value in shared_strings:
        strings += record(19, rich_string(value) if isinstance(value, str) else value)
    strings += record(160)
    rels.append('<Relationship Id="strings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.bin"/>')
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="bin" ContentType="' + MAIN_TYPE + '"/>' + ''.join(f'<Override PartName="/xl/worksheets/sheet{index}.bin" ContentType="application/vnd.ms-excel.worksheet"/>' for index in range(1, len(sheets) + 1)) + '</Types>')
        archive.writestr("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="book" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.bin"/></Relationships>')
        archive.writestr("xl/workbook.bin", book)
        archive.writestr("xl/_rels/workbook.bin.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + ''.join(rels) + '</Relationships>')
        archive.writestr("xl/styles.bin", styles)
        archive.writestr("xl/sharedStrings.bin", strings)
        for index, (_, contents) in enumerate(sheets, 1):
            archive.writestr(f"xl/worksheets/sheet{index}.bin", contents)
    return output.getvalue()


def replace_part(data, path, contents):
    output = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as original, zipfile.ZipFile(output, "w") as changed:
        for item in original.infolist():
            changed.writestr(item, contents if item.filename == path else original.read(item.filename))
    return output.getvalue()


def records(data):
    """Decode independently so tests do not reuse the production parser."""
    offset = 0
    while offset < len(data):
        start = offset
        kind = data[offset]
        offset += 1
        if kind & 128:
            kind = (kind & 127) | (data[offset] << 7)
            offset += 1
        size = 0
        for shift in range(4):
            part = data[offset]
            offset += 1
            size |= (part & 127) << (7 * shift)
            if not part & 128:
                break
        end = offset + size
        if end > len(data):
            raise ValueError("truncated record")
        yield kind, data[offset:end], start, end
        offset = end


def recompute_index(sheet):
    """Encode Excel binaryIndex blocks independently from the writer under test."""
    data_types = set(range(1, 12)) | {62}
    rows, row_offsets = {}, {}
    begin = row = None
    for kind, payload, start, end in records(sheet):
        if kind == 145:
            begin = start
        elif kind == 0:
            row = struct.unpack_from("<I", payload)[0]
            row_offsets[row] = start
        elif kind in data_types:
            col = struct.unpack_from("<I", payload)[0]
            rows.setdefault(row, []).append((col, start))
    starts = []
    for row in sorted(rows):
        if not starts or row >= starts[-1] + 32:
            starts.append(row)
    output = b""
    if not starts:
        output = record(42, struct.pack("<IIII", 0, 32, 0, 0) + bytes(8))
    for index, first in enumerate(starts):
        last = min(first + 32, 1048576)
        data_rows = [row for row in range(first, last) if row in rows]
        base = begin if index == 0 else row_offsets[data_rows[0]]
        mask, column_masks, offsets = 0, [], []
        for row in data_rows:
            mask |= 1 << (row - first)
            column_mask, seen = 0, set()
            for col, offset in rows[row]:
                block = col // 1024
                if block not in seen:
                    seen.add(block)
                    column_mask |= 1 << block
                    offsets.append((row, block, offset - base))
            column_masks.append(column_mask)
        row_block = struct.pack("<IQ", mask, base)
        row_block += b"".join(struct.pack("<H", value) for value in column_masks)
        row_block += b"".join(struct.pack("<I", value) for _, _, value in sorted(offsets))
        unused = bytes(((last - first + 32) // 32) * 4)
        payload = struct.pack("<IIII", first, last, 0, 0) + unused
        following = len(output) + len(record(42, payload))
        output += record(42, struct.pack("<IIII", first, last, following, 0) + unused)
        output += record(40, row_block)
    return output + record(277)


def with_indexes(data, indexes=None, *, calc_chain=False, extra_parts=None):
    """Add index relationships and package parts for rewrite boundary tests."""
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    if indexes is None:
        indexes = {}
        for path, contents in parts.items():
            if path.startswith("xl/worksheets/sheet") and path.endswith(".bin"):
                index = int(path.rsplit("sheet", 1)[1][:-4])
                indexes[index] = recompute_index(contents)
    content_types = parts["[Content_Types].xml"].decode("utf-8")
    for index, contents in indexes.items():
        path = f"xl/worksheets/binaryIndex{index}.bin"
        parts[path] = contents
        relpath = f"xl/worksheets/_rels/sheet{index}.bin.rels"
        parts[relpath] = (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="index" Type="http://schemas.microsoft.com/office/2006/relationships/xlBinaryIndex" Target="binaryIndex{index}.bin"/>'
            '</Relationships>').encode("utf-8")
        content_types = content_types.replace('</Types>', f'<Override PartName="/{path}" ContentType="application/vnd.ms-excel.binIndexWs"/></Types>')
    if calc_chain:
        parts["xl/calcChain.bin"] = record(528)
        parts["xl/_rels/calcChain.bin.rels"] = b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
        relpath = "xl/_rels/workbook.bin.rels"
        rels = parts[relpath].decode("utf-8").replace('</Relationships>',
            '<Relationship Id="calc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.bin"/></Relationships>')
        parts[relpath] = rels.encode("utf-8")
        content_types = content_types.replace('</Types>', '<Override PartName="/xl/calcChain.bin" ContentType="application/vnd.ms-excel.calcChain"/></Types>')
    parts["[Content_Types].xml"] = content_types.encode("utf-8")
    parts.update(extra_parts or {})
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path, contents in parts.items():
            archive.writestr(path, contents)
    return output.getvalue()
